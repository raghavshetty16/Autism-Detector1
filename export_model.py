from pathlib import Path

import joblib
import pandas as pd
from sklearn.metrics import make_scorer, fbeta_score
from sklearn.model_selection import GridSearchCV, train_test_split
from sklearn.preprocessing import MinMaxScaler
from sklearn.svm import SVC


BASE_DIR = Path(__file__).resolve().parent
DATA_PATH = BASE_DIR / "autism_data.csv"
MODEL_DIR = BASE_DIR / "models"
MODEL_PATH = MODEL_DIR / "autism_full_model.joblib"
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
NUMERIC_FEATURES = ["age", "result"]
CATEGORICAL_FEATURES = [
    "gender",
    "ethnicity",
    "jundice",
    "austim",
    "contry_of_res",
    "relation",
]


def f_beta_score(y_true, y_predict):
    return fbeta_score(y_true, y_predict, beta=0.5)


def main():
    data = pd.read_csv(DATA_PATH).dropna(subset=RAW_FEATURES + ["Class/ASD"])

    features_raw = data[RAW_FEATURES]

    scaler = MinMaxScaler()
    features_minmax_transform = pd.DataFrame(data=features_raw)
    features_minmax_transform[NUMERIC_FEATURES] = scaler.fit_transform(
        features_raw[NUMERIC_FEATURES]
    )

    x = pd.get_dummies(features_minmax_transform)
    y = data["Class/ASD"].apply(lambda value: 1 if value == "YES" else 0)

    x_train, _, y_train, _ = train_test_split(
        x,
        y,
        test_size=0.2,
        random_state=1,
    )

    parameters = {
        "C": range(1, 6),
        "kernel": ["linear", "poly", "rbf", "sigmoid"],
        "degree": range(1, 6),
    }
    grid_obj = GridSearchCV(
        estimator=SVC(random_state=1),
        param_grid=parameters,
        scoring=make_scorer(f_beta_score),
    )
    grid_fit = grid_obj.fit(x_train.values, y_train)

    MODEL_DIR.mkdir(exist_ok=True)
    joblib.dump(
        {
            "model": grid_fit.best_estimator_,
            "scaler": scaler,
            "raw_features": RAW_FEATURES,
            "numeric_features": NUMERIC_FEATURES,
            "categorical_features": CATEGORICAL_FEATURES,
            "encoded_columns": list(x.columns),
            "category_options": {
                column: sorted(data[column].astype(str).unique())
                for column in CATEGORICAL_FEATURES
            },
            "model_name": "SVC GridSearchCV from autism_detection.ipynb full features",
        },
        MODEL_PATH,
    )
    print(f"Saved model to {MODEL_PATH}")
    print(f"Best parameters: {grid_fit.best_params_}")


if __name__ == "__main__":
    main()
