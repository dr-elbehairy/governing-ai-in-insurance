"""Step 1 — load freMPL1 (CASdatasets export) and prepare the analysis frame.
Input : data/freMPL1.csv     Output: data/freMPL1_prepared.csv, results_data.json"""
import os, json, pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
df = pd.read_csv(f"{HERE}/data/freMPL1.csv")
df = df[df.Exposure > 0].copy()
df["SocioGrp"] = df["SocioCateg"].str[:4]          # 37 codes -> 6 major CSP groups
df["VehAge"] = df["VehAge"].astype(str)
df["Female"] = (df.Gender == "Female").astype(int)
df["freq"] = df.ClaimInd / df.Exposure
keep = ["Exposure","ClaimInd","freq","Gender","Female","VehAge","MariStat","SocioGrp","VehUsage","VehBody","VehPrice",
        "VehEngine","VehEnergy","VehMaxSpeed","VehClass","Garage","LicAge","DrivAge","HasKmLimit","BonusMalus","RiskVar"]
df[keep].to_csv(f"{HERE}/data/freMPL1_prepared.csv", index=False)
raw = df.groupby("Gender").apply(lambda g: g.ClaimInd.sum() / g.Exposure.sum())
meta = {"n_policies": int(len(df)), "claim_rate": float(df.ClaimInd.mean()), "female_share": float(df.Female.mean()),
        "raw_freq": {k: float(v) for k, v in raw.items()}, "n_rating_factors": 16}
json.dump(meta, open(f"{HERE}/results_data.json", "w"), indent=2); print(meta)
