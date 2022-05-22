import openpyxl
import numpy as np
from kivy.app import App
from kivy.lang import Builder
from kivy.properties import StringProperty
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.config import Config
import kivy.utils
from kivy.uix.scrollview import ScrollView
from kivy.uix.popup import Popup

Config.set('graphics', 'width', '600')
Config.set('graphics', 'height', '900')





class WindowManager(ScreenManager):
    pass

class MainWindow(Screen):
    pass


class DiagnoseWindow(Screen):
    wb = openpyxl.load_workbook("Data.xlsx")

    def importData(sheet):
        return ([[cell.value for cell in row] for row in sheet])

    # import Properties
    Properties = wb['Properties']
    P = importData(Properties['A2:E48'])
    # import base
    fuzzyData = wb['fuzzyData']
    base = importData(fuzzyData['A1:N199'])
    # import C
    sheetC = wb['C']
    C = importData(sheetC['A1:AFZ199'])

    def fuzzy(self, Xi, input):

        if Xi == 13:
            if input.upper() == "TSG NẶNG":
                return 2
            if input.upper() == "TSG":
                return 1
            if input.upper() == "THAI BÌNH THƯỜNG":
                return 0
        else:
            for X in self.P:
                if Xi != X[0]:
                    continue
                else:
                    if input is None:
                        return None
                    if float(input) >= float(X[2]) and float(input) <= float(X[3]):
                        return X[4]
                    else:
                        continue

    def fuzzyList(self, X):
        list = []
        for i in range(len(X)):
            list.append(self.fuzzy(i, X[i]))

        return list


    def combination(self, k, n):
        if k == 0 or k == n:
            return 1
        if k == 1:
            return n
        return self.combination(k - 1, n - 1) + self.combination(k, n - 1)


    def FISA(self, base, C, list):
        colum = len(base[0])
        row = len(base)

        cols = self.combination(3, (colum - 1))
        C0 = [0] * cols
        C1 = [0] * cols
        C2 = [0] * cols

        t = 0
        for a in range(0, colum - 3):
            for b in range(a + 1, colum - 2):
                for c in range(b + 1, colum - 1):
                    for r in range(row):
                        if base[r][a] == list[a] and base[r][b] == list[b] and base[r][c] == list[c] and base[r][
                            colum - 1] == 0:
                            C0[t] = C[r][t + 0 * cols]
                            break
                        elif base[r][a] == list[a] and base[r][b] == list[b] and base[r][c] == list[c] and base[r][
                            colum - 1] == 1:
                            C1[t] = C[r][t + 1 * cols]
                            break
                        elif base[r][a] == list[a] and base[r][b] == list[b] and base[r][c] == list[c] and base[r][
                            colum - 1] == 2:
                            C2[t] = C[r][t + 2 * cols]
                            break
                    t += 1

        D0 = max(C0) + min(C0)
        D1 = max(C1) + min(C1)
        D2 = max(C2) + min(C2)

        # print(D0)
        # print(D1)
        # print(D2)
        if D0 > D1 and D0 > D2:
            return 0
        elif D1 > D2:
            return 1
        else:
            return 2


    def diagnose(self, x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13):
        X = [x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13]
        list = self.fuzzyList(X)
        result = self.FISA(self.base, self.C, list)

        return result


    def on_click(self, x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11, x12, x13):
        r = self.diagnose(float(x1), float(x2), float(x3), float(x4), float(x5), float(x6), float(x7), float(x8), float(x9), float(x10), float(x11), float(x12), float(x13))
        if r == 0:
            ResultPopup.diagnoseResult = "Chúc mừng, bạn đang hoàn toàn khỏe mạnh !"
        elif r == 1:
            ResultPopup.diagnoseResult = "Bạn đang có nguy cơ mắc tiền sản giật"
        elif r == 2:
            ResultPopup.diagnoseResult = "Bạn đang có nguy cơ mắc tiền sản giật nặng"


class ScrollViewDiagnose(ScrollView):
    pass


class ResultPopup(Popup):
    diagnoseResult = StringProperty("")
    pass


class tutorialPopup(Popup):
    pass


class ScrollViewTutorial(ScrollView):
    pass


class FKGApplication(App):
    def build(self):
        return kv


kv = Builder.load_file("TSG.kv")

if __name__ == "__main__":
    FKGApplication().run()