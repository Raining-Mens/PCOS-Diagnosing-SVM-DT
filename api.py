from flask import Flask, request, jsonify, redirect, url_for, render_template
import pickle
import os
import pandas as pd
from openpyxl.workbook import Workbook
from openpyxl import load_workbook

from werkzeug.utils import secure_filename

app = Flask(__name__)
model = pickle.load(open('models\svm-model.pkl', 'rb'))


@app.route("/")
def home():
    return render_template("index.html")


app.config["EXCEL_UPLOADS"] = "./static/assets/uploads"
app.config["ALLOWED_EXCEL_EXTENSIONS"] = ["XLSX", "CSV", "XLS"]


def allowed_excel(filename):

    if not "." in filename:
        return False

    ext = filename.rsplit(".", 1)[1]

    if ext.upper() in app.config["ALLOWED_EXCEL_EXTENSIONS"]:
        return True
    else:
        return False


@app.route("/tool", methods=["GET", "POST"])
def tool():


    if request.method == "POST":

        if request.files:

            excel = request.files["input"]

            if excel.filename == "":
                print("Excel file must have a filename")
                return redirect(request.url)

            if not allowed_excel(excel.filename):
                print("That excel extension is not allowed")
                return redirect(request.url)

            else:
                filename = secure_filename(excel.filename)

                excel.save(os.path.join(app.config["EXCEL_UPLOADS"], filename))

            print("excel save")

            wb = load_workbook(os.path.join(app.config["EXCEL_UPLOADS"], filename))

            ws = wb.active

            print(ws["D26"].value)

            CycleRI = ws["D26"].value
            FSHmIUmL = ws["E26"].value
            LHmIUmL = ws["F26"].value
            AMHngmL = ws["G26"].value
            PulseRateBPM = ws["H26"].value
            PRGngmL = ws["I26"].value
            RBSmgdl = ws["J26"].value
            BP_SystolicmmHg = ws["K26"].value
            BP_DiastolicmmHg = ws["L26"].value
            AvgFsizeLmm = ws["M26"].value
            AvgFsizeRmm = ws["N26"].value
            Endometriummm = ws["O26"].value
            Age = ws["P26"].value
            Hairgrowth = ws["Q26"].value
            SkinDarkening = ws["R26"].value

            makeprediction = model.predict([[CycleRI, FSHmIUmL, LHmIUmL,
                                            AMHngmL, PulseRateBPM, PRGngmL, RBSmgdl,
                                            BP_SystolicmmHg, BP_DiastolicmmHg, AvgFsizeLmm, AvgFsizeRmm,
                                            Endometriummm, Age, Hairgrowth, SkinDarkening]])

            output = round(makeprediction[0], 2)

            print(output)

            return redirect(request.url)
        
    return render_template("tool.html")


@app.route("/predict", methods=["GET", "POST"])
def predict():
    
    CycleRI = request.args.get('Cycle(R/I)')
    FSHmIUmL = request.args.get('FSH(mIU/mL)')
    LHmIUmL = request.args.get('LH(mIU/mL)')
    AMHngmL = request.args.get('AMH(ng/mL)')
    PulseRateBPM = request.args.get('Pulse rate(bpm)')
    PRGngmL = request.args.get('PRG(ng/mL)')
    RBSmgdl = request.args.get('RBS(mg/dl)')
    BP_SystolicmmHg = request.args.get('BP _Systolic (mmHg)')
    BP_DiastolicmmHg = request.args.get('BP _Diastolic (mmHg)')
    AvgFsizeLmm = request.args.get('Avg. F size (L) (mm)')
    AvgFsizeRmm = request.args.get('Avg. F size (R) (mm)')
    Endometriummm = request.args.get('Endometrium (mm)')
    Age = request.args.get('Age (yrs)')
    Hairgrowth = request.args.get('hair growth(Y/N)')
    SkinDarkening = request.args.get('Skin darkening (Y/N)')

    makeprediction = model.predict([[CycleRI, FSHmIUmL, LHmIUmL,
                                     AMHngmL, PulseRateBPM, PRGngmL, RBSmgdl,
                                     BP_SystolicmmHg, BP_DiastolicmmHg, AvgFsizeLmm, AvgFsizeRmm,
                                     Endometriummm, Age, Hairgrowth, SkinDarkening]])

    output = round(makeprediction[0], 2)

    return jsonify({'PCOS:': output})


if __name__ == "__main__":
    app.run(debug=True)
