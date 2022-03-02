import csv
import pandas as pd
import numpy as np

# read file of excel
fileExcel = "Modulo 3. BS MedCar2020_Aplicativo.xlsm"


# read first sheet
def creaateDetalles(fileExcel):
    dfDetalles = pd.read_excel(fileExcel, sheet_name="Detalles")
    with open('DetallesBS.csv', 'w', newline='', encoding='utf-8') as f:
        # create header for csv
        header = [dfDetalles.iloc[7][1], dfDetalles.iloc[11][1], dfDetalles.iloc[12][1], dfDetalles.iloc[13][1],
                  dfDetalles.iloc[17][1], dfDetalles.iloc[18][1], dfDetalles.iloc[19][1], dfDetalles.iloc[20][1],
                  dfDetalles.iloc[23][1], dfDetalles.iloc[24][1], dfDetalles.iloc[25][1], dfDetalles.iloc[8][8],
                  dfDetalles.iloc[9][8], dfDetalles.iloc[10][8], dfDetalles.iloc[11][8], dfDetalles.iloc[12][8],
                  dfDetalles.iloc[13][8], dfDetalles.iloc[14][8], dfDetalles.iloc[15][8], dfDetalles.iloc[11][13],
                  dfDetalles.iloc[12][13], dfDetalles.iloc[13][13], dfDetalles.iloc[19][8], dfDetalles.iloc[18][10],
                  dfDetalles.iloc[18][11], dfDetalles.iloc[17][6], dfDetalles.iloc[10][13]]
        fileCsv = csv.DictWriter(f, fieldnames=header)
        fileCsv.writeheader()
        # write row whit values
        fileCsv.writerow({header[0]: dfDetalles.iloc[7][2], header[1]: dfDetalles.iloc[11][4],
                          header[2]: dfDetalles.iloc[12][4], header[3]: dfDetalles.iloc[13][4],
                          header[4]: dfDetalles.iloc[17][5], header[5]: dfDetalles.iloc[18][5],
                          header[6]: dfDetalles.iloc[19][5], header[7]: dfDetalles.iloc[20][5],
                          header[8]: dfDetalles.iloc[23][5], header[9]: dfDetalles.iloc[24][5],
                          header[10]: dfDetalles.iloc[25][5], header[11]: dfDetalles.iloc[8][11],
                          header[12]: dfDetalles.iloc[9][11], header[13]: dfDetalles.iloc[10][11],
                          header[14]: dfDetalles.iloc[11][11], header[15]: dfDetalles.iloc[12][11],
                          header[16]: dfDetalles.iloc[13][11], header[17]: dfDetalles.iloc[14][11],
                          header[18]: dfDetalles.iloc[15][11], header[19]: dfDetalles.iloc[11][13],
                          header[20]: dfDetalles.iloc[12][13], header[21]: dfDetalles.iloc[13][13],
                          header[22]: dfDetalles.iloc[19][10], header[23]: dfDetalles.iloc[19][11],
                          header[24]: dfDetalles.iloc[19][12], header[25]: dfDetalles.iloc[17][6],
                          header[26]: dfDetalles.iloc[10][13]})


def createAjustePresion(fileExcel):
    test = pd.read_excel(fileExcel, sheet_name="Ajuste Presión")
    test = test.iloc[5:, :17]
    test.to_csv('Ajuste PresiónBS.csv', header=False, index=False)


def createBsReportados(fileExcel):
    test = pd.read_excel(fileExcel, sheet_name="BSreportados")
    test = test.iloc[5:, :34]
    test.to_csv('BsReportados.csv', header=False, index=False)


def createWeldToWeld(fileExcel):
    test = pd.read_excel(fileExcel, sheet_name="WeldToWeld")
    test = test.iloc[5:, :18]
    test.iloc[0, 3] = "Soldadura emparejada [Si / No]"
    test.iloc[0, 7] = "Diferencia en deformación total"
    test.iloc[0, 8] = "TStrain 2020 - 2018"
    test.iloc[0, 12] = "Soldadura emparejada[Si / No]"
    test.iloc[0, 16] = "Diferencia en deformación total"
    test.iloc[0, 17] = "TStrain 2020 - 2011"
    test.to_csv('WeldToWeld.csv', header=False, index=False)


def createXYZdeltaP(fileExcel):
    test = pd.read_excel(fileExcel, sheet_name="XYZdeltaP")
    test = test.iloc[5:, :47]
    test.to_csv('XYZdeltaP.csv', header=False, index=False)


def createXYZdeltaOp1(fileExcel):
    test = pd.read_excel(fileExcel, sheet_name="XYZdeltaOp1")
    test = test.iloc[5:, :13]
    test.to_csv('XYZdeltaOp1.csv', header=False, index=False)


def createXYZdeltaOp2(fileExcel):
    test = pd.read_excel(fileExcel, sheet_name="XYZdeltaOp2")
    test = test.iloc[5:, :13]
    test.to_csv('XYZdeltaOp2.csv', header=False, index=False)


# this function return the index of number less in list
def match(valor, list):
    cont = 0
    for i in list:
        if valor > i:
            cont = cont + 1
    return cont - 1


# interpolation line
def interpolate(value, x, y):
    x = np.array(x)
    y = np.array(y)
    value = np.array(value)
    N = len(x)
    A = np.zeros((N, N))
    b = y
    for i in range(N):
        A.T[i] = pow(x, i)
    a = np.linalg.solve(A, b)
    valueY = a[0] + a[1] * value
    return valueY


# this function fill sheet "AjustePresionBS.csv"
def extractData(AjustePresiónBS, Detalles):
    df = pd.read_csv(AjustePresiónBS)
    df1 = pd.read_csv(Detalles)
    DF = df1.iloc[0, 18]  # column 18 in DetallesBS.csv (Factor de diseño)
    range1 = df.iloc[0:, 0]  # 'Pressure Sheet- Design Pressure
    for row, value in enumerate(range1):
        SMYS = df.iloc[row, 2]  # column 2, smys Psi
        De = df.iloc[row, 3] / 25.4  # column 3, Diámetro [mm]
        t = df.iloc[row, 4] / 25.4  # column 4, Espesor [mm]
        df.iloc[row, 12] = np.ceil(2 * DF * t * SMYS / De)  # column 13, Presión de diseño [psi]

    # 'Forecast MOP by Weld
    datacountW = len(range1)
    rdregW = range1

    df.iloc[0, 16] = np.minimum(df.iloc[0, 0], df.iloc[0, 14])  # value min in column 0 and 14;row 0
    drStart = df.iloc[1, 14]  # column 14,(PK)
    nh = len(df.iloc[0:, 14])  # column 14,(PK), size
    drAbsolute = df.iloc[1:, 14]  # all dates in column 14,(PK del perfil hidraúlico) start in 1
    for row, dr in enumerate(drAbsolute):
        df.iloc[row + 1, 16] = np.round(((dr - drStart) * 1000), 2)

    drAbsolute = df.iloc[0:nh, 16]  # all dates in column 16,(Distancia de registro perfil operativo [m])
    if rdregW.min() < drAbsolute.min():
        print("El primer valor de distancia de registro del perfíl debe ser menor al DR de \
        Assessment Fuera de intervalo")
    elif rdregW.max() > drAbsolute.max() or datacountW < 0:
        print("El perfil no corresponde, porque no cubre todas las distancias de registro establecidas para cada \
        número de soldadura. Es necesario aumentar el PK final para que cubra todas las distancias de registro\
        Valores incorrectos")

    dReg = [drAbsolute[0], drAbsolute[1]]
    pPresiones = df.iloc[0:nh, 15]  # column 15, (Perfil de presión operativa [psi])
    ph = [pPresiones[0], pPresiones[1]]
    listdrAbsolute = drAbsolute.tolist()

    for row, dist in enumerate(rdregW):
        df.iloc[row, 11] = df.iloc[row, 0] / 100  # column 11, (Distancia de Registro [km])
        if dist > dReg[0]:
            index = match(dist, listdrAbsolute)
            dReg = [drAbsolute[index], drAbsolute[index + 1]]
            ph = [pPresiones[index], pPresiones[index + 1]]
        df.iloc[row, 13] = np.ceil(interpolate(dist, dReg, ph))  # column 13,(Presión Interpolada [psi])

    df.to_csv('Ajuste PresiónBS.csv', index=False)


def beding_multianual(df):
    lista = []  # create list for columns
    for i in range(0, 25, 1):
        lista.append(i)
    print(lista)
    df = df.iloc[:, lista]  # select column 0 to 24

    print(df)


# this function is for validate dates, option => distancia de registro(1) / soldadura(0)
def validateBsReportados(AjustePresionBS, BsReportados, option):
    df = pd.read_csv(AjustePresionBS)
    df1 = pd.read_csv(BsReportados)
    if option:
        rDreg = df1.iloc[0:, 1]  # column 1 (Distancia de registro [m]) in BsReportados.csv
        rDregP = df.iloc[0:, 0]  # column 0 (Distancia absoluta [m]) in AjustePresionBS.csv
        dataCount = np.maximum(len(df1.iloc[0:, 1]), len(df1.iloc[0:, 0]))  # size max of column 0 and 1

        if dataCount > 0 and len(
                rDregP) > 0:  # size dates in colum 0 or 1(BsReportados.csv), column 0(AjustePresion.csv)
            for row, rdreg in enumerate(rDreg):
                # if is a number and if not is null and match is number
                if rDreg[row].dtype == 'float64' and not pd.isna(rdreg) and match(rdreg, rDregP.tolist()):
                    df1.iloc[row, 14] = df.iloc[match(rdreg, rDregP.tolist()), 2]
                    if df1.iloc[row, 14] <= 31000 and df1.iloc[row, 14] >= 29500:
                        df1.iloc[row, 15] = 48600
                    elif df1.iloc[row, 14] <= 36000 and df1.iloc[row, 14] >= 34000:
                        df1.iloc[row, 15] = 60200
                    elif df1.iloc[row, 14] <= 43000 and df1.iloc[row, 14] >= 41000:
                        df1.iloc[row, 15] = 60200
                    elif df1.iloc[row, 14] <= 47000 and df1.iloc[row, 14] >= 45000:
                        df1.iloc[row, 15] = 63100
                    elif df1.iloc[row, 14] <= 53000 and df1.iloc[row, 14] >= 51000:
                        df1.iloc[row, 15] = 66700
                    elif df1.iloc[row, 14] <= 57000 and df1.iloc[row, 14] >= 54000:
                        df1.iloc[row, 15] = 71100
                    elif df1.iloc[row, 14] <= 62000 and df1.iloc[row, 14] >= 59000:
                        df1.iloc[row, 15] = 75400
                    elif df1.iloc[row, 14] <= 67000 and df1.iloc[row, 14] >= 64000:
                        df1.iloc[row, 15] = 77600
                    elif df1.iloc[row, 14] <= 71000 and df1.iloc[row, 14] >= 69000:
                        df1.iloc[row, 15] = 82700
                    else:
                        df1.iloc[row, 15] = ""
                    # -----------------CheckBox2
                    df1.iloc[row, 16] = df.iloc[match(rdreg, rDregP.tolist()), 3]  # column 16 (Diametro)
                    # -----------------CheckBox4
                    df1.iloc[row, 13] = df.iloc[match(rdreg, rDregP.tolist()), 4]  # column 13 (Espesor de pared)
                    # -----------------CheckBox3
                    df1.iloc[row, 18] = df.iloc[match(rdreg, rDregP.tolist()), 13]  # column 18 (Presion Interna)
                    df1.iloc[row, 19] = df1.iloc[row, 18] / 2  # column 19 (Presion interna al 50%)
                else:
                    print(
                        "Valor erroneo de distancia de registro, es necesario que sea un valor númerico. Dato incorrecto")

    else:
        rdregWeld = df1.iloc[0:, 11]  # column 11 (Número de soldadura)
        count = len(df.iloc[0:, 0])
        print(rdregWeld[0].dtype)
        if len(rdregWeld) > 0:
            for row, value in enumerate(rdregWeld):
                # if is number and is not null
                if rdregWeld[row] == 'float64' or rdregWeld[row] == 'float' or rdregWeld[row] == 'int64' \
                        or not pd.isna(value):
                    range1 = df.iloc[0:, 0:count]

    beding_multianual(df1)
    df1.to_csv('BsReportados.csv', index=False)


creaateDetalles(fileExcel)
createAjustePresion(fileExcel)
createBsReportados(fileExcel)
# createWeldToWeld(fileExcel)
# createXYZdeltaP(fileExcel)
# createXYZdeltaOp1(fileExcel)
# createXYZdeltaOp2(fileExcel)
extractData('Ajuste PresiónBS.csv', 'DetallesBS.csv')
validateBsReportados('Ajuste PresiónBS.csv', "BsReportados.csv", 1)
