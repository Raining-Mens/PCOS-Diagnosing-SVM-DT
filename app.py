from flask import Flask, request, jsonify, redirect, url_for, render_template, session
import pickle
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
from werkzeug.utils import secure_filename
import os
import git
from predict_funct import *

app = Flask(__name__)

app.secret_key = "hello"

THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
app.config.from_object("config")
app.config["EXCEL_UPLOADS"] = "static/assets/uploads"
app.config["ASSETS"] = "static/assets"
app.config["ALLOWED_EXCEL_EXTENSIONS"] = ["XLSX", "CSV", "XLS"]
my_excel = os.path.join(THIS_FOLDER, "static/assets/uploads")
my_assets = os.path.join(THIS_FOLDER, "static/assets")


@app.route('/git_update', methods=['POST'])
def git_update():
    repo = git.Repo('./PCOS-Diagnosing-SVM-DT')
    origin = repo.remotes.origin
    repo.create_head('main',
                     origin.refs.main).set_tracking_branch(origin.refs.main).checkout()
    origin.pull()
    return '', 200


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/ovarian_index")
def ovarian_index():
    return render_template("ovarian-index.html")


@app.route("/pcos_index")
def pcos_index():
    return render_template("pcos-index.html")


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
    session.pop("result", None)
    session.pop("model", None)
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
                excel.save(os.path.join(my_excel, filename))
                session['save_excel'] = filename

            output = predict_excel_svm(excel)
            session['result'] = int(output)

            return redirect(url_for("result"))
    else:
        if "result" in session:
            return redirect(url_for("pop"))
    return render_template("tool.html")


@app.route("/dt", methods=["GET", "POST"])
def dt():
    session.pop("result", None)
    session.pop("model", None)
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
                excel.save(os.path.join(my_excel, filename))
                session['save_excel'] = filename

            output = predict_excel_dt(excel)
            session['result'] = int(output)

            return redirect(url_for("result"))
    else:
        if "result" in session:
            return redirect(url_for("pop"))
    return render_template("dt.html")


@app.route("/ovariansvm", methods=["GET", "POST"])
def ovariansvm():
    session.pop("result", None)
    session.pop("model", None)
    session.pop("PatID", None)
    session.pop("Age", None)
    session.pop("Hairgrowth", None)
    session.pop("CycleRI", None)
    session.pop("WeightGain", None)
    session.pop("FastFood", None)

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
                excel.save(os.path.join(my_excel, filename))
                session['save_excel'] = filename

            output = ovarian_svm(excel)
            session['result'] = int(output)

            return redirect(url_for("ovarianresult"))
    else:
        if "result" in session:
            return redirect(url_for("pop"))
    return render_template("ovariansvm.html")


@app.route("/ovariandt", methods=["GET", "POST"])
def ovariandt():
    session.pop("result", None)
    session.pop("model", None)
    session.pop("PatID", None)
    session.pop("Age", None)
    session.pop("Hairgrowth", None)
    session.pop("CycleRI", None)
    session.pop("WeightGain", None)
    session.pop("FastFood", None)
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
                excel.save(os.path.join(my_excel, filename))
                session['save_excel'] = filename

            output = ovarian_dt(excel)
            session['result'] = int(output)

            return redirect(url_for("ovarianresult"))
    else:
        if "result" in session:
            return redirect(url_for("pop"))
    return render_template("ovariandt.html")


@app.route("/result", methods=["GET", "POST"])
def result():
    save_excel = session['save_excel']
    book = load_workbook(open(os.path.join(my_excel, save_excel), 'rb'))
    sheet = book.active

    if "result" in session:
        result = session["result"]
        model = session["model"]
        PatID = session["PatID"]
        Age = session["Age"]
        Hairgrowth = session["Hairgrowth"]
        CycleRI = session["CycleRI"]
        WeightGain = session["WeightGain"]
        FastFood = session["FastFood"]
        # Errorint = session['errorint']
        # print(Errorint)
        if model == "SVM":
            model_name = "SVM"
        else:
            model_name = "DT"

        print(result)
        if result == 1:
            return render_template("results.html", RESULTS="POSITIVE", EXCEL=sheet, MODEL=model_name, ID=PatID, AGE=Age,
                                   HAIR=Hairgrowth, CYC=CycleRI, AFL=WeightGain, AFR=FastFood)
        elif result == 0:
            return render_template("results.html", RESULTS="NEGATIVE", EXCEL=sheet, MODEL=model_name, ID=PatID, AGE=Age,
                                   HAIR=Hairgrowth, CYC=CycleRI, AFL=WeightGain, AFR=FastFood)
    else:
        return redirect(url_for("tool"))


@app.route("/ovarianresult", methods=["GET", "POST"])
def ovarianresult():
    save_excel = session['save_excel']
    book = load_workbook(open(os.path.join(my_excel, save_excel), 'rb'))
    sheet = book.active

    if "result" in session:
        result = session["result"]
        model = session["model"]

        Menopause = session["Menopause"]
        CANine = session["CANine"]
        CASeven = session["CASeven"]
        AeFP = session["AeFP"]
        CAOneTwo = session["CAOneTwo"]
        Age = session["Age"]
        if model == "SVM":
            model_name = "SVM"
        else:
            model_name = "DT"

        print(result)
        if result == 0:
            return render_template("ovarian-result.html", RESULTS="POSITIVE", EXCEL=sheet, MODEL=model_name, AGE=Age,
                                   HAIR=Menopause, CYC=CANine, AFL=CASeven, AFR=AeFP, CAT=CAOneTwo)
        elif result == 1:
            return render_template("ovarian-result.html", RESULTS="NEGATIVE", EXCEL=sheet, MODEL=model_name, AGE=Age,
                                   HAIR=Menopause, CYC=CANine, AFL=CASeven, AFR=AeFP, CAT=CAOneTwo)
    else:
        return redirect(url_for("tool"))


@app.route("/about")
def about():
    return render_template("about.html")


@app.route("/pop")
def pop():
    session.pop("result", None)
    session.pop("model", None)
    session.pop("result", None)
    session.pop("model", None)
    session.pop("PatID", None)
    session.pop("Age", None)
    session.pop("Hairgrowth", None)
    session.pop("CycleRI", None)
    session.pop("WeightGain", None)
    session.pop("FastFood", None)
    return redirect(url_for("tool"))


if __name__ == "__main__":
    app.run(debug=True)
