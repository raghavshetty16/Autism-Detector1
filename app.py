from pathlib import Path
from urllib.parse import parse_qs

import joblib
import pandas as pd
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field, ValidationError, field_validator

from database import save_user_response


BASE_DIR = Path(__file__).resolve().parent
MODEL_PATH = BASE_DIR / "models" / "autism_full_model.joblib"
AQ10_FEATURES = [f"A{i}_Score" for i in range(1, 11)]
RAW_FEATURES = [
    "age",
    "gender",
    "ethnicity",
    "jundice",
    "austim",
    "contry_of_res",
    "result",
    "relation",
    *AQ10_FEATURES,
]

QUESTIONS = [
    {
        "name": "A1_Score",
        "number": "1",
        "text": "I often notice small sounds when others do not.",
        "yes_scores_one": True,
    },
    {
        "name": "A2_Score",
        "number": "2",
        "text": "I usually concentrate more on the whole picture, rather than the small details.",
        "yes_scores_one": False,
    },
    {
        "name": "A3_Score",
        "number": "3",
        "text": "I find it easy to do more than one thing at once.",
        "yes_scores_one": False,
    },
    {
        "name": "A4_Score",
        "number": "4",
        "text": "If there is an interruption, I can switch back to what I was doing very quickly.",
        "yes_scores_one": False,
    },
    {
        "name": "A5_Score",
        "number": "5",
        "text": "I find it easy to read between the lines when someone is talking to me.",
        "yes_scores_one": False,
    },
    {
        "name": "A6_Score",
        "number": "6",
        "text": "I know how to tell if someone listening to me is getting bored.",
        "yes_scores_one": False,
    },
    {
        "name": "A7_Score",
        "number": "7",
        "text": "When I am reading a story, I find it difficult to work out the characters' intentions.",
        "yes_scores_one": True,
    },
    {
        "name": "A8_Score",
        "number": "8",
        "text": "I like to collect information about categories of things.",
        "yes_scores_one": True,
    },
    {
        "name": "A9_Score",
        "number": "9",
        "text": "I find it easy to work out what someone is thinking or feeling just by looking at their face.",
        "yes_scores_one": False,
    },
    {
        "name": "A10_Score",
        "number": "10",
        "text": "I find it difficult to work out people's intentions.",
        "yes_scores_one": True,
    },
]
QUESTIONS_BY_NAME = {question["name"]: question for question in QUESTIONS}
PROFILE_LABELS = {
    "age": "Age",
    "gender": "Gender",
    "ethnicity": "Ethnicity",
    "jundice": "Jaundice history",
    "austim": "Family autism history",
    "contry_of_res": "Country of residence",
    "result": "AQ-10 total score",
    "relation": "Completed by",
}


def load_model_bundle() -> dict:
    if not MODEL_PATH.exists():
        raise RuntimeError(
            "Model file not found. Run the model export cell in autism_detection.ipynb "
            "or run `python export_model.py` once before starting FastAPI."
        )
    return joblib.load(MODEL_PATH)


class ScreeningAnswers(BaseModel):
    age: float = Field(..., ge=1, le=120)
    gender: str
    ethnicity: str
    jundice: str
    austim: str
    contry_of_res: str
    relation: str
    A1_Score: int = Field(..., ge=0, le=1)
    A2_Score: int = Field(..., ge=0, le=1)
    A3_Score: int = Field(..., ge=0, le=1)
    A4_Score: int = Field(..., ge=0, le=1)
    A5_Score: int = Field(..., ge=0, le=1)
    A6_Score: int = Field(..., ge=0, le=1)
    A7_Score: int = Field(..., ge=0, le=1)
    A8_Score: int = Field(..., ge=0, le=1)
    A9_Score: int = Field(..., ge=0, le=1)
    A10_Score: int = Field(..., ge=0, le=1)
    result: float | None = None

    @field_validator(*AQ10_FEATURES, mode="before")
    @classmethod
    def normalize_yes_no(cls, value: object) -> int:
        if isinstance(value, str):
            cleaned = value.strip().lower()
            if cleaned in {"yes", "y", "true", "1"}:
                return 1
            if cleaned in {"no", "n", "false", "0"}:
                return 0
        return int(value)


app = FastAPI(title="Autism AQ-10 Screening")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")
model_bundle = load_model_bundle()
model = model_bundle["model"]
scaler = model_bundle["scaler"]
raw_features = model_bundle["raw_features"]
numeric_features = model_bundle["numeric_features"]
categorical_features = model_bundle["categorical_features"]
encoded_columns = model_bundle["encoded_columns"]
category_options = model_bundle["category_options"]


def format_profile_value(feature: str, value: object) -> str:
    if value == "?":
        return "Unknown"
    if feature == "gender":
        return "Female" if value == "f" else "Male" if value == "m" else str(value)
    if feature in {"jundice", "austim"}:
        return str(value).capitalize()
    if feature == "age":
        return str(int(value)) if float(value).is_integer() else str(value)
    return str(value)


def describe_encoded_feature(column: str) -> str:
    if column in PROFILE_LABELS:
        return PROFILE_LABELS[column]

    if column in QUESTIONS_BY_NAME:
        return f"Answer to question {QUESTIONS_BY_NAME[column]['number']}"

    for feature in categorical_features:
        prefix = f"{feature}_"
        if column.startswith(prefix):
            value = column[len(prefix):]
            return f"{PROFILE_LABELS.get(feature, feature)}: {format_profile_value(feature, value)}"

    return column


def build_aq_explanation(
    row: dict,
    raw_question_answers: dict[str, str] | None = None,
) -> list[dict]:
    items = []
    for question in QUESTIONS:
        feature = question["name"]
        scored = int(row[feature])
        raw_answer = raw_question_answers.get(feature) if raw_question_answers else None
        scoring_answer = "Yes" if question["yes_scores_one"] else "No"

        items.append(
            {
                "number": question["number"],
                "question": question["text"],
                "answer": raw_answer.capitalize() if raw_answer else f"Score {scored}",
                "scored": scored,
                "status": "1 point" if scored else "0 points",
                "rule": f"A {scoring_answer} answer counts as 1 point.",
            }
        )
    return items


def build_profile_summary(row: dict) -> list[dict]:
    return [
        {
            "label": PROFILE_LABELS[feature],
            "value": format_profile_value(feature, row[feature]),
        }
        for feature in raw_features
        if feature not in AQ10_FEATURES
    ]


def build_model_factor_explanation(encoded_frame: pd.DataFrame) -> dict:
    if not hasattr(model, "coef_"):
        return {
            "available": False,
            "supporting": [],
            "opposing": [],
            "note": "This model does not expose linear feature weights.",
        }

    coefficients = model.coef_[0]
    values = encoded_frame.iloc[0].to_numpy()
    contributions = []

    for column, value, coefficient in zip(encoded_columns, values, coefficients):
        contribution = float(value * coefficient)
        if abs(contribution) < 0.0001:
            continue
        contributions.append(
            {
                "feature": describe_encoded_feature(column),
                "impact": contribution,
                "strength": round(abs(contribution), 3),
                "direction": "YES" if contribution > 0 else "NO",
            }
        )

    yes_factors = sorted(
        [item for item in contributions if item["impact"] > 0],
        key=lambda item: item["impact"],
        reverse=True,
    )[:5]
    no_factors = sorted(
        [item for item in contributions if item["impact"] < 0],
        key=lambda item: item["impact"],
    )[:5]

    return {
        "available": True,
        "supporting_yes": yes_factors,
        "supporting_no": no_factors,
        "note": "Feature factors come from the linear SVC decision weights for this saved model.",
    }


def predict_from_answers(
    answers: ScreeningAnswers,
    raw_question_answers: dict[str, str] | None = None,
) -> dict:
    row = answers.model_dump()
    row["result"] = sum(row[feature] for feature in AQ10_FEATURES)

    raw_frame = pd.DataFrame([[row[feature] for feature in raw_features]], columns=raw_features)
    transformed_frame = raw_frame.copy()
    transformed_frame[numeric_features] = scaler.transform(raw_frame[numeric_features])
    encoded_frame = pd.get_dummies(transformed_frame)
    encoded_frame = encoded_frame.reindex(columns=encoded_columns, fill_value=0)

    prediction = int(model.predict(encoded_frame.values)[0])

    probability = None
    if hasattr(model, "predict_proba"):
        probability = round(float(model.predict_proba(encoded_frame.values)[0][1]) * 100, 2)

    model_factors = build_model_factor_explanation(encoded_frame)
    main_factors = (
        model_factors["supporting_yes"]
        if prediction == 1
        else model_factors["supporting_no"]
    )
    opposing_factors = (
        model_factors["supporting_no"]
        if prediction == 1
        else model_factors["supporting_yes"]
    )

    return {
        "prediction": "YES" if prediction == 1 else "NO",
        "probability": probability,
        "score": int(row["result"]),
        "answers": {feature: row[feature] for feature in raw_features},
        "explanation": {
            "summary": (
                "Your answers and details were closer to the YES examples the model learned from."
                if prediction == 1
                else "Your answers and details were closer to the NO examples the model learned from."
            ),
            "aq_items": build_aq_explanation(row, raw_question_answers),
            "profile": build_profile_summary(row),
            "main_factors": main_factors,
            "opposing_factors": opposing_factors,
            "model_note": model_factors["note"],
        },
    }


def score_form_answers(raw_answers: dict[str, str]) -> dict[str, int]:
    scored_answers = {}
    for feature, value in raw_answers.items():
        cleaned = value.strip().lower()
        if cleaned not in {"yes", "no"}:
            raise ValueError(f"{feature} must be answered as yes or no.")

        yes_scores_one = QUESTIONS_BY_NAME[feature]["yes_scores_one"]
        scored_answers[feature] = int(
            (cleaned == "yes" and yes_scores_one)
            or (cleaned == "no" and not yes_scores_one)
        )
    return scored_answers


def build_saved_response(
    answers: ScreeningAnswers,
    raw_question_answers: dict[str, str] | None,
    result: dict,
    request: Request | None = None,
) -> dict:
    payload = answers.model_dump()
    payload["result"] = result["score"]

    document = {
        "profile": {
            feature: payload[feature]
            for feature in raw_features
            if feature not in AQ10_FEATURES and feature != "result"
        },
        "aq10_scored_answers": {
            feature: payload[feature]
            for feature in AQ10_FEATURES
        },
        "aq10_raw_answers": raw_question_answers,
        "model_output": {
            "prediction": result["prediction"],
            "probability": result["probability"],
            "score": result["score"],
            "model_name": model_bundle.get("model_name"),
            "explanation": result.get("explanation"),
        },
    }

    if request is not None:
        document["request"] = {
            "client_host": request.client.host if request.client else None,
            "user_agent": request.headers.get("user-agent"),
        }

    return document


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "questions": QUESTIONS,
            "category_options": category_options,
            "result": None,
            "form_values": {},
            "error": None,
        },
    )


@app.post("/predict")
async def predict_form(request: Request):
    body = (await request.body()).decode("utf-8")
    form_data = parse_qs(body)
    raw_answers = {
        feature: form_data.get(feature, ["no"])[0]
        for feature in AQ10_FEATURES
    }
    profile_answers = {
        feature: form_data.get(feature, [""])[0]
        for feature in raw_features
        if feature not in AQ10_FEATURES and feature != "result"
    }

    try:
        scored_answers = score_form_answers(raw_answers)
        answers = ScreeningAnswers(**profile_answers, **scored_answers)
    except (ValueError, ValidationError):
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "questions": QUESTIONS,
                "category_options": category_options,
                "result": None,
                "form_values": {**profile_answers, **raw_answers},
                "error": "Please complete every field before predicting.",
            },
            status_code=400,
        )

    result = predict_from_answers(answers, raw_answers)
    saved_to_mongodb = save_user_response(
        build_saved_response(answers, raw_answers, result, request)
    )

    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "questions": QUESTIONS,
            "category_options": category_options,
            "result": result,
            "form_values": {**profile_answers, **raw_answers},
            "saved_to_mongodb": saved_to_mongodb,
            "error": None,
        },
    )


@app.post("/api/predict")
async def predict_api(answers: ScreeningAnswers):
    try:
        result = predict_from_answers(answers)
        saved_to_mongodb = save_user_response(
            build_saved_response(answers, None, result)
        )
        return {**result, "saved_to_mongodb": saved_to_mongodb}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/assets/image.jpeg")
async def hero_image():
    return FileResponse(BASE_DIR / "image.jpeg")
