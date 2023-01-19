from flask import Flask, request, jsonify, redirect, url_for, render_template, session
import pickle
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
from werkzeug.utils import secure_filename
import os
import git

THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))
EXCEL_UPLOADS = "static/assets/uploads"
ASSETS = "static/assets"
my_excel = os.path.join(THIS_FOLDER, "static/assets/uploads")
my_assets = os.path.join(THIS_FOLDER, "static/assets")


def predict_excel_svm(excel):
    wb = load_workbook(excel)

    ws = wb.active

    PatID = ws["A2"].value
    Age = ws["B2"].value
    Hairgrowth = ws["I2"].value
    SkinDarkening = ws["J2"].value
    PulseRateBPM = ws["Q2"].value
    CycleRI = ws["T2"].value
    FSHmIUmL = ws["AA2"].value
    LHmIUmL = ws["AB2"].value
    AMHngmL = ws["AE2"].value
    PRGngmL = ws["AH2"].value
    RBSmgdl = ws["AI2"].value
    BP_SystolicmmHg = ws["AJ2"].value
    BP_DiastolicmmHg = ws["AK2"].value
    AvgFsizeLmm = ws["AN2"].value
    AvgFsizeRmm = ws["AO2"].value
    Endometriummm = ws["AP2"].value

    session["PatID"] = PatID
    session["Age"] = Age
    session["Hairgrowth"] = Hairgrowth
    session["CycleRI"] = CycleRI
    session["AvgFsizeLmm"] = AvgFsizeLmm
    session["AvgFsizeRmm"] = AvgFsizeRmm


    model = pickle.load(open(os.path.join(my_assets, "svm-model.pkl"), 'rb'))
    session['model'] = "SVM"


    makeprediction = model.predict([[Age, Hairgrowth, SkinDarkening,
                                    PulseRateBPM, CycleRI, FSHmIUmL, LHmIUmL,
                                    AMHngmL, PRGngmL, RBSmgdl, BP_SystolicmmHg,
                                    BP_DiastolicmmHg, AvgFsizeLmm, AvgFsizeRmm, Endometriummm]])

    output = round(makeprediction[0], 2)

    return(output)

def predict_excel_dt(excel):
    wb = load_workbook(excel)

    ws = wb.active

    PatID = ws["A2"].value
    Age = ws["B2"].value
    Hairgrowth = ws["I2"].value
    SkinDarkening = ws["J2"].value
    PulseRateBPM = ws["Q2"].value
    CycleRI = ws["T2"].value
    FSHmIUmL = ws["AA2"].value
    LHmIUmL = ws["AB2"].value
    AMHngmL = ws["AE2"].value
    PRGngmL = ws["AH2"].value
    RBSmgdl = ws["AI2"].value
    BP_SystolicmmHg = ws["AJ2"].value
    BP_DiastolicmmHg = ws["AK2"].value
    AvgFsizeLmm = ws["AN2"].value
    AvgFsizeRmm = ws["AO2"].value
    Endometriummm = ws["AP2"].value

    session["PatID"] = PatID
    session["Age"] = Age
    session["Hairgrowth"] = Hairgrowth
    session["CycleRI"] = CycleRI
    session["AvgFsizeLmm"] = AvgFsizeLmm
    session["AvgFsizeRmm"] = AvgFsizeRmm


    model = pickle.load(open(os.path.join(my_assets, "dt-model.pkl"), 'rb'))
    session['model'] = "DT"


    makeprediction = model.predict([[Age, Hairgrowth, SkinDarkening,
                                    PulseRateBPM, CycleRI, FSHmIUmL, LHmIUmL,
                                    AMHngmL, PRGngmL, RBSmgdl, BP_SystolicmmHg,
                                    BP_DiastolicmmHg, AvgFsizeLmm, AvgFsizeRmm, Endometriummm]])

    output = round(makeprediction[0], 2)

    return(output)


def ovarian_svm(excel):
    wb = load_workbook(excel)

    ws = wb.active

    Age = ws["C2"].value
    Menopause = ws["AH2"].value
    CANine = ws["M2"].value
    CASeven = ws["N2"].value
    AeFP = ws["A2"].value
    CAOneTwo = ws["L2"].value
    HEFour = ws["Z2"].value
    CEyA = ws["O2"].value


    session["Age"] = Age
    session["Menopause"] = Menopause
    session["CANine"] = CANine
    session["CASeven"] = CASeven
    session["AeFP"] = AeFP
    session["CAOneTwo"] = CAOneTwo


    model = pickle.load(open(os.path.join(my_assets, "svm-ovarian.pkl"), 'rb'))
    session['model'] = "SVM"


    makeprediction = model.predict([[Age, Menopause, CANine,
                                    CASeven, AeFP, CAOneTwo, HEFour,
                                    CEyA]])

    output = round(makeprediction[0], 2)

    return(output)


def ovarian_dt(excel):
    wb = load_workbook(excel)

    ws = wb.active

    Age = ws["C2"].value
    Menopause = ws["AH2"].value
    CANine = ws["M2"].value
    CASeven = ws["N2"].value
    AeFP = ws["A2"].value
    CAOneTwo = ws["L2"].value
    HEFour = ws["Z2"].value
    CEyA = ws["O2"].value


    session["Age"] = Age
    session["Menopause"] = Menopause
    session["CANine"] = CANine
    session["CASeven"] = CASeven
    session["AeFP"] = AeFP
    session["CAOneTwo"] = CAOneTwo


    model = pickle.load(open(os.path.join(my_assets, "dt-ovarian.pkl"), 'rb'))
    session['model'] = "DT"


    makeprediction = model.predict([[Age, Menopause, CANine,
                                    CASeven, AeFP, CAOneTwo, HEFour,
                                    CEyA]])

    output = round(makeprediction[0], 2)

    return(output)