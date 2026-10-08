import os
from datetime import datetime, timezone
from typing import Any

from dotenv import load_dotenv
from pymongo import MongoClient, ReplaceOne
from pymongo.errors import PyMongoError


load_dotenv()

MONGODB_DB_NAME = os.getenv("MONGODB_DB_NAME", "autism_detection")
TRAINING_COLLECTION = os.getenv("MONGODB_TRAINING_COLLECTION", "training_data")
RESPONSES_COLLECTION = os.getenv("MONGODB_RESPONSES_COLLECTION", "user_responses")

_client: MongoClient | None = None


def get_mongo_client() -> MongoClient | None:
    global _client

    uri = os.getenv("MONGODB_URI")
    if not uri:
        return None

    if _client is None:
        _client = MongoClient(uri, serverSelectionTimeoutMS=3000)
    return _client


def get_database():
    client = get_mongo_client()
    if client is None:
        return None
    return client[MONGODB_DB_NAME]


def save_user_response(document: dict[str, Any]) -> bool:
    db = get_database()
    if db is None:
        return False

    document = {
        **document,
        "created_at": datetime.now(timezone.utc),
        "source": "fastapi_frontend",
    }

    try:
        db[RESPONSES_COLLECTION].insert_one(document)
    except PyMongoError:
        return False

    return True


def upsert_training_rows(rows: list[dict[str, Any]]) -> int:
    db = get_database()
    if db is None:
        raise RuntimeError("MONGODB_URI is not configured.")

    operations = [
        ReplaceOne({"dataset_row_id": row["dataset_row_id"]}, row, upsert=True)
        for row in rows
    ]
    if not operations:
        return 0

    result = db[TRAINING_COLLECTION].bulk_write(operations)
    return result.upserted_count + result.modified_count
