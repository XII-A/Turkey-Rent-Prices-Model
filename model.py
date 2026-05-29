"""Benchmark models for residential rent price prediction in Turkiye.

The raw structured export is `final_data_v4.csv`. The reproducible modeling
dataset used here is `finalDataModel.csv`, which contains the numeric features
created during preprocessing.
"""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
import tensorflow as tf
from sklearn.linear_model import LinearRegression
from sklearn.metrics import (
    explained_variance_score,
    mean_absolute_error,
    mean_squared_error,
    r2_score,
)
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import MinMaxScaler
from sklearn.tree import DecisionTreeRegressor
from tensorflow.keras import regularizers
from tensorflow.keras.callbacks import EarlyStopping
from tensorflow.keras.layers import BatchNormalization, Dense, Dropout, Input
from tensorflow.keras.models import Sequential
from tensorflow.keras.optimizers import Adam


RAW_DATA_PATH = Path("final_data_v4.csv")
MODEL_DATA_PATH = Path("finalDataModel.csv")
METRICS_PATH = Path("model_metrics.json")
RANDOM_STATE = 101
TEST_SIZE = 0.30
MAX_EPOCHS = 400
BATCH_SIZE = 128


def get_z_score(values: pd.Series) -> pd.Series:
    return (values - values.mean()) / values.std()


def load_model_data(path: Path = MODEL_DATA_PATH) -> tuple[pd.DataFrame, pd.Series]:
    df = pd.read_csv(path, low_memory=False, index_col=False).dropna().copy()

    df["log_fiyat"] = np.log1p(df["fiyat"])
    df = df.drop(columns=["fiyat"])
    df = df[get_z_score(df["log_fiyat"]) < 3].copy()

    df["rooms"] = df["living_rooms"] + df["bedrooms"]
    df = df.drop(columns=["living_rooms", "bedrooms"])

    X = df.drop(columns=["log_fiyat"])
    y = df["log_fiyat"]
    return X, y


def evaluate_model(name: str, model, X_train, X_test, y_train, y_test) -> dict:
    model.fit(X_train, y_train)
    predictions = model.predict(X_test)
    mse = mean_squared_error(y_test, predictions)

    return {
        "model": name,
        "mse": float(mse),
        "mae": float(mean_absolute_error(y_test, predictions)),
        "rmse": float(np.sqrt(mse)),
        "r2": float(r2_score(y_test, predictions)),
        "explained_variance": float(explained_variance_score(y_test, predictions)),
    }


def build_ann_model(input_shape: int) -> Sequential:
    model = Sequential(
        [
            Input(shape=(input_shape,)),
            Dense(64, activation="relu"),
            Dense(64, activation="relu", kernel_regularizer=regularizers.l2(0.01)),
            BatchNormalization(),
            Dropout(0.1),
            Dense(64, activation="relu", kernel_regularizer=regularizers.l2(0.01)),
            BatchNormalization(),
            Dropout(0.1),
            Dense(64, activation="relu", kernel_regularizer=regularizers.l2(0.01)),
            Dense(1),
        ]
    )
    model.compile(optimizer=Adam(learning_rate=0.001), loss="mse", metrics=["mse"])
    return model


def evaluate_ann(X_train, X_test, y_train, y_test) -> dict:
    tf.keras.utils.set_random_seed(RANDOM_STATE)
    model = build_ann_model(X_train.shape[1])
    early_stopping = EarlyStopping(
        monitor="val_loss",
        patience=10,
        restore_best_weights=True,
    )

    history = model.fit(
        X_train,
        y_train.to_numpy(),
        validation_data=(X_test, y_test.to_numpy()),
        batch_size=BATCH_SIZE,
        epochs=MAX_EPOCHS,
        callbacks=[early_stopping],
        verbose=0,
    )

    predictions = model.predict(X_test, verbose=0).reshape(-1)
    mse = mean_squared_error(y_test, predictions)

    return {
        "model": "Artificial Neural Network (Keras)",
        "mse": float(mse),
        "mae": float(mean_absolute_error(y_test, predictions)),
        "rmse": float(np.sqrt(mse)),
        "r2": float(r2_score(y_test, predictions)),
        "explained_variance": float(explained_variance_score(y_test, predictions)),
        "epochs_trained": int(len(history.history["loss"])),
    }


def main() -> None:
    X, y = load_model_data()

    X_train, X_test, y_train, y_test = train_test_split(
        X,
        y,
        test_size=TEST_SIZE,
        random_state=RANDOM_STATE,
    )

    scaler = MinMaxScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_test_scaled = scaler.transform(X_test)

    results = [
        evaluate_model(
            "Linear Regression",
            LinearRegression(),
            X_train_scaled,
            X_test_scaled,
            y_train,
            y_test,
        ),
        evaluate_model(
            "Decision Tree",
            DecisionTreeRegressor(random_state=RANDOM_STATE, min_samples_leaf=10),
            X_train_scaled,
            X_test_scaled,
            y_train,
            y_test,
        ),
        evaluate_ann(X_train_scaled, X_test_scaled, y_train, y_test),
    ]

    summary = {
        "raw_source_dataset": str(RAW_DATA_PATH),
        "modeling_dataset": str(MODEL_DATA_PATH),
        "target": "log1p(fiyat)",
        "rows_after_cleaning": int(len(y)),
        "feature_count": int(X.shape[1]),
        "test_size": TEST_SIZE,
        "random_state": RANDOM_STATE,
        "results": results,
    }

    METRICS_PATH.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
