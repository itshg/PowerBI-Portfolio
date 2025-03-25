from flask import Flask, request, jsonify, render_template
import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor
from sklearn.cluster import KMeans
from office365.sharepoint.client_context import ClientContext
import os

app = Flask(__name__)

class PowerBIAI:
    def __init__(self, file=None, sharepoint_url=None, client_id=None, client_secret=None):
        self.file = file
        self.sharepoint_url = sharepoint_url
        self.client_id = client_id
        self.client_secret = client_secret
        self.data = self.load_data()
    
    def load_data(self):
        if self.file:
            return pd.read_csv(self.file) if self.file.filename.endswith('.csv') else pd.read_excel(self.file)
        elif self.sharepoint_url:
            return self.load_from_sharepoint()
        return None
    
    def load_from_sharepoint(self):
        ctx = ClientContext(self.sharepoint_url).with_credentials(self.client_id, self.client_secret)
        lists = ctx.web.lists.get_by_title("Leads").get().execute_query()
        items = lists.items.get().execute_query()
        data = [{field: item.properties[field] for field in item.properties} for item in items]
        return pd.DataFrame(data)
    
    def preprocess_data(self):
        self.data.fillna(self.data.mean(), inplace=True)
        scaler = StandardScaler()
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        self.data[numeric_cols] = scaler.fit_transform(self.data[numeric_cols])
    
    def generate_insights(self):
        insights = {
            "Total Leads": len(self.data),
            "Average Revenue per Lead": self.data["Estimated Revenue"].mean() if "Estimated Revenue" in self.data else None,
            "Conversion Rate": self.data["Converted"].mean() if "Converted" in self.data else None,
            "Revenue Forecast": self.revenue_forecasting(),
            "Customer Segments": self.customer_segmentation()
        }
        return insights

    def revenue_forecasting(self):
        if "Estimated Revenue" in self.data.columns and "Lead Date" in self.data.columns:
            self.data["Lead Date"] = pd.to_datetime(self.data["Lead Date"])
            self.data.sort_values(by="Lead Date", inplace=True)
            self.data["Days"] = (self.data["Lead Date"] - self.data["Lead Date"].min()).dt.days
            X = self.data[["Days"]]
            y = self.data["Estimated Revenue"]
            model = RandomForestRegressor(n_estimators=100)
            model.fit(X, y)
            future_days = np.array([[X.max() + i] for i in range(1, 31)])
            predictions = model.predict(future_days)
            return predictions.mean()
        return None
    
    def customer_segmentation(self):
        if "Estimated Revenue" in self.data.columns:
            X = self.data[["Estimated Revenue"]]
            kmeans = KMeans(n_clusters=3, random_state=42)
            self.data["Segment"] = kmeans.fit_predict(X)
            return self.data["Segment"].value_counts().to_dict()
        return None
    
    def generate_powerbi_dashboard(self):
        return "https://powerbi.com/report/generated_dashboard"

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"})
    file = request.files['file']
    ai = PowerBIAI(file=file)
    ai.preprocess_data()
    insights = ai.generate_insights()
    insights["Power BI Dashboard"] = ai.generate_powerbi_dashboard()
    return jsonify(insights)

if __name__ == '__main__':
    app.run(debug=True)
