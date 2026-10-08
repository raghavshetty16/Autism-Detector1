# Autism Detection Using Machine Learning

A FastAPI web application that uses an AQ-10 questionnaire and profile details to screen for autism-related traits. The application is based on the experiments in `autism_detection.ipynb`, but it uses a saved model for predictions so the model is not trained every time the server starts.

> **Important:** This project provides a machine-learning screening result. It is not a medical diagnosis and must not replace an assessment by a qualified healthcare professional.

![Autism screening application](image.jpeg)

## What the Application Does

1. The user enters basic profile information.
2. The user answers 10 AQ-10 questions with **Yes** or **No**.
3. The application converts the answers into AQ-10 points.
4. The saved machine-learning model receives the AQ-10 points and profile details.
5. The result appears in a modal as **YES** or **NO**.
6. The modal explains which answers and details influenced the result.
7. When MongoDB is configured, the response is saved for possible future model training.

The interface also preserves the submitted answers, allowing the user to close the result modal and make changes.

## Main Features

- Real AQ-10 questions instead of labels such as `A1` and `A2`
- Consistent Yes/No controls and responsive styling
- FastAPI backend with a Jinja2 frontend
- Pre-trained model loaded from a `joblib` file
- No model training during application startup or prediction
- Full feature preprocessing used by the exported notebook model
- Plain-language result explanation
- AQ-10 point-by-point answer breakdown
- MongoDB Atlas storage for completed responses
- JSON prediction API for programmatic use
- Responsive result modal for desktop and mobile screens

## Technology Stack

| Area | Technology |
| --- | --- |
| Backend | FastAPI |
| Server | Uvicorn |
| Frontend | HTML, CSS, JavaScript, Jinja2 |
| Data processing | pandas |
| Machine learning | scikit-learn |
| Saved model | joblib |
| Database | MongoDB Atlas with PyMongo |
| Configuration | python-dotenv |

## Project Structure

```text
.
|-- app.py                         # FastAPI routes, validation and prediction logic
|-- database.py                    # MongoDB connection and storage helpers
|-- export_model.py                # One-time model training and export script
|-- autism_detection.ipynb         # Original data analysis and model experiments
|-- autism_data.csv                # Original training dataset
|-- image.jpeg                     # Image used by the frontend and this README
|-- models/
|   `-- autism_full_model.joblib   # Saved model and preprocessing information
|-- templates/
|   `-- index.html                 # Screening form and result modal
|-- static/
|   `-- styles.css                 # Application styling and responsive rules
|-- requirements.txt               # Python dependencies
|-- .env.example                   # Example MongoDB configuration
|-- .env                           # Local secrets; ignored by Git
`-- .gitignore                     # Files that must not be committed
```

## How the Model Was Created

### Notebook experiments

The notebook explores five machine-learning techniques:

1. Decision Tree
2. Random Forest
3. Support Vector Machine (SVM)
4. K-Nearest Neighbors (KNN)
5. Logistic Regression

These techniques are used during experimentation to compare how different classifiers perform on the dataset. Accuracy and the F-beta score are used to evaluate predictions. The F-beta calculation uses `beta=0.5`, which gives more weight to precision than recall.

### Final exported model

The production model is created by `export_model.py`. It:

1. Reads `autism_data.csv`.
2. Removes rows missing required model values.
3. scales `age` and the AQ-10 total score with `MinMaxScaler`.
4. converts text categories into numeric one-hot columns with `pandas.get_dummies()`.
5. splits the data into 80% training data and 20% test data using `random_state=1`.
6. runs `GridSearchCV` over several SVC settings.
7. selects the SVC configuration with the best F-beta score.
8. saves the chosen model and all preprocessing information to `models/autism_full_model.joblib`.

The saved model currently contains a linear SVC selected with `C=1` and `degree=1`. The bundle also contains the fitted scaler, expected input columns, one-hot encoded columns, category choices and model name.

This saving process is usually called **model serialization** or **model persistence**. Loading the saved file later is called **deserialization**.

## Features Used for Prediction

The exported model uses 18 raw input columns:

| Group | Fields |
| --- | --- |
| Numeric | `age`, `result` |
| Profile categories | `gender`, `ethnicity`, `jundice`, `austim`, `contry_of_res`, `relation` |
| AQ-10 scores | `A1_Score` through `A10_Score` |

`result` is not entered manually. The application calculates it by adding the 10 scored AQ answers.

Some field names such as `jundice`, `austim`, and `contry_of_res` contain spelling mistakes because those are the names used by the original dataset and trained model. The frontend shows readable labels such as **Jaundice history**, **Family autism history**, and **Country of residence**.

After one-hot encoding, the saved model expects 103 numeric columns. The application recreates those columns in the same order before making a prediction.

## AQ-10 Scoring

Every question produces either `1` point or `0` points. A Yes answer does not always mean one point because some AQ-10 statements are reverse-scored.

| Question | Answer that adds 1 point |
| --- | --- |
| 1 | Yes |
| 2 | No |
| 3 | No |
| 4 | No |
| 5 | No |
| 6 | No |
| 7 | Yes |
| 8 | Yes |
| 9 | No |
| 10 | Yes |

For example, question 2 says that the person usually concentrates on the whole picture. A **No** answer adds one point because this question is reverse-scored.

## Prediction Flow

```text
Browser form
    |
    v
FastAPI validates profile fields and Yes/No answers
    |
    v
AQ-10 answers are converted to 0 or 1 points
    |
    v
The 10 points are added to calculate the AQ-10 total
    |
    v
Age and total score are scaled
    |
    v
Category values are one-hot encoded
    |
    v
Columns are matched to the saved model's 103-column order
    |
    v
The saved SVC predicts YES or NO
    |
    +--> Result and explanation are shown in the modal
    |
    `--> Response is saved to MongoDB when available
```

The model is loaded once when `app.py` starts. A prediction request does not train the five notebook models and does not run `GridSearchCV` again.

## Result Explanation

Because the saved SVC uses a linear kernel, the application can inspect how each encoded input contributes to its decision. The modal displays:

- the prediction and AQ-10 total score;
- the submitted profile details;
- details that supported the final result;
- details that pointed in the opposite direction; and
- how every AQ-10 answer received zero or one point.

These contributions explain the saved model's calculation. They do not prove why a person does or does not have autism.

## MongoDB Storage

MongoDB is optional. Predictions still work when MongoDB is unavailable, but the response will not be saved.

The default database is `autism_detection` and it contains two collections:

| Collection | Purpose |
| --- | --- |
| `training_data` | Original dataset rows uploaded for reference or later training |
| `user_responses` | New responses submitted through the frontend or API |

A saved user response contains:

- profile values;
- raw Yes/No answers when submitted through the browser;
- scored AQ-10 answers;
- prediction, score, confidence when available, and explanation;
- model name;
- UTC creation time and source; and
- browser request metadata for frontend submissions.

Saving responses does **not** automatically retrain or replace the model. Before future responses are used for supervised training, they should be reviewed, cleaned, and given a reliable outcome label. Retraining should then be run deliberately, evaluated, and exported as a new model file.

## Installation

### 1. Open the project directory

```powershell
cd C:\path\to\Autism_Detection_using_Machine_Learning-main
```

### 2. Create a virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, run this command for the current terminal and try again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

### 3. Install dependencies

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Configure MongoDB

Create `.env` from `.env.example` and enter the MongoDB Atlas connection string:

```env
MONGODB_URI=mongodb+srv://<username>:<password>@<cluster-url>/?appName=Cluster0
MONGODB_DB_NAME=autism_detection
MONGODB_TRAINING_COLLECTION=training_data
MONGODB_RESPONSES_COLLECTION=user_responses
```

Never commit `.env`. It contains database credentials and is already excluded by `.gitignore`.

For MongoDB Atlas, also ensure that:

- the database user has permission to read and write;
- the current IP address is allowed under **Network Access**; and
- special characters in the username or password are URL-encoded.

### 5. Confirm that the model exists

The application expects:

```text
models/autism_full_model.joblib
```

The repository already contains this file. Only regenerate it when intentionally retraining the model:

```powershell
python export_model.py
```

### 6. Start the application

```powershell
python -m uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

Open [http://127.0.0.1:8000](http://127.0.0.1:8000) in a browser.

If port 8000 is blocked or already in use, choose another port:

```powershell
python -m uvicorn app:app --reload --host 127.0.0.1 --port 8010
```

## API Usage

The JSON endpoint is:

```text
POST /api/predict
```

For the API, `A1_Score` through `A10_Score` should normally be sent as already-scored `0` or `1` values. The browser form performs the reverse scoring automatically before calling the prediction code.

Example request:

```powershell
$body = @{
    age = 24
    gender = "m"
    ethnicity = "White-European"
    jundice = "no"
    austim = "no"
    contry_of_res = "India"
    relation = "Self"
    A1_Score = 1
    A2_Score = 0
    A3_Score = 1
    A4_Score = 0
    A5_Score = 1
    A6_Score = 0
    A7_Score = 1
    A8_Score = 0
    A9_Score = 1
    A10_Score = 0
} | ConvertTo-Json

Invoke-RestMethod `
    -Uri "http://127.0.0.1:8000/api/predict" `
    -Method Post `
    -ContentType "application/json" `
    -Body $body
```

The response contains the prediction, calculated AQ-10 score, submitted model values, explanation and a `saved_to_mongodb` flag.

Interactive API documentation is available while the server is running:

- Swagger UI: [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs)
- ReDoc: [http://127.0.0.1:8000/redoc](http://127.0.0.1:8000/redoc)

## Common Problems

### Socket permission error

If Uvicorn reports `WinError 10013`, use a different host and port:

```powershell
python -m uvicorn app:app --reload --host 127.0.0.1 --port 8010
```

Also check whether Windows Firewall, antivirus software, or another process is blocking the selected port.

### Model file not found

Run:

```powershell
python export_model.py
```

Then confirm that `models/autism_full_model.joblib` was created.

### Responses are not saved

Check the `.env` values, Atlas database-user permissions and Atlas Network Access rules. The frontend intentionally continues showing a prediction even if saving fails.

### Changes do not appear in the browser

Use `Ctrl+F5` to perform a hard refresh. The `--reload` option restarts FastAPI after Python file changes, but the browser may still cache CSS.

## Privacy and Responsible Use

This application can collect sensitive health-related answers and request metadata. Before public deployment:

- obtain clear user consent;
- explain what is stored and why;
- collect only the data that is genuinely needed;
- restrict MongoDB access;
- rotate exposed credentials immediately;
- use HTTPS;
- define a retention and deletion policy;
- avoid treating predictions as confirmed diagnoses; and
- have a qualified professional review the wording and workflow.

## Development Notes

- Keep the model bundle and preprocessing rules together. Changing preprocessing without retraining can make predictions unreliable.
- Do not rename model input columns unless the model is retrained and exported with matching names.
- Test both YES and NO result paths after changing the modal.
- Test with MongoDB available and unavailable because saving is designed to fail gracefully.
- Evaluate a newly trained model on data it did not see during training before replacing the production `joblib` file.

## Disclaimer

The AQ-10 is a screening questionnaire, and this project adds a machine-learning prediction trained on a limited dataset. The result can be affected by the data quality, the dataset population, preprocessing choices and model limitations. A healthcare professional should perform any diagnosis or clinical evaluation.
