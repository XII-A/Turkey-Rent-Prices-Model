# Data-Driven Analysis and Prediction of Residential Rent Prices in Turkiye

This project uses machine learning to estimate continuous residential rent
prices in Turkiye from public rental listing attributes.

## Project workflow

1. `webscraper.py` collects public rental listing data from Emlakjet.
2. `preprocessing.ipynb` cleans and transforms the raw listing data.
3. `final_data_v4.csv` stores the richer structured source export.
4. `finalDataModel.csv` stores the processed numeric modeling dataset.
5. `model.py` trains reproducible benchmark regressors, including the Keras
   artificial neural network, and records metrics in `model_metrics.json`.

## Features

The model-ready dataset includes rental price, net/gross area, floor and
building information, heating type, site/balcony status, bathroom counts,
room counts, balcony-type indicators, latitude, and longitude.

## Current benchmark results

The target variable is `log1p(fiyat)`. After missing-value removal, outlier
filtering, room-count feature engineering, MinMax scaling, and a 70/30 train
test split, the current benchmark results are:

| Model | MSE | MAE | RMSE | R2 |
| --- | ---: | ---: | ---: | ---: |
| Linear Regression | 0.3609 | 0.3758 | 0.6007 | 0.2443 |
| Decision Tree | 0.2764 | 0.3074 | 0.5257 | 0.4212 |
| Artificial Neural Network (Keras) | 0.3120 | 0.3605 | 0.5586 | 0.3466 |

The original project report noted an ANN MSE of `0.22`. The table above shows
the fresh reproducible Keras run from the current script and environment, so it
may differ from the earlier experiment.

## Run benchmarks

```powershell
python model.py
```

The script writes the latest metrics to `model_metrics.json`.
