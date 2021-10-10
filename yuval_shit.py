from openpyxl import load_workbook
import pandas as pd


class TimeSaver:

    def __init__(self, gdud_path, shihva):
        self.gdud_path = gdud_path
        self.shihva_ascii = ord(shihva)

    def save_my_time(self):
        splited_names = self.get_hanihim()
        reshumim = self.search_in_excel(splited_names)
        self.save_excel_file(reshumim)

    def get_hanihim(self):
        global madrih
        f = open(self.gdud_path, "r", encoding='utf-8')
        lines = f.readlines()
        f.close()

        splited_names = {}
        for name in lines:
            if name == "\n":
                continue
            if ":\n" in name:
                madrih = name.replace(":\n", "")
                splited_names[madrih] = []
            else:
                splited_name = name[:-1].split(" ", 1)
                splited_name[0] = splited_name[0].replace("-", " ")
                splited_names[madrih].append(splited_name)
        return splited_names

    def search_in_excel(self, splited_names, path="data.xlsx"):
        workbook = load_workbook(filename=path)
        sheet = workbook.active
        index = 1
        reshumim = []
        while sheet[f"A{index}"].value:
            try:
                parrant, phone_number = str(sheet[f"R{index}"].value).split("-")
            except:
                parrant = "לא נמצא"
                phone_number = sheet[f"R{index}"].value
            current_name = [sheet[f"E{index}"].value, sheet[f"D{index}"].value, "אין", "לא", parrant,
                            phone_number, "אין", sheet[f"F{index}"].value, sheet[f"S{index}"].value]

            for m in splited_names.keys():
                if current_name[:2] in splited_names[m] and \
                        ord(sheet[f"K{index}"].value[0]) == self.shihva_ascii:
                    current_name.insert(2, m)
                    reshumim.append(current_name)

            index += 1

        f = open("full_shit.txt", "w", encoding='utf-8')
        index = 1
        for i in reshumim:
            s = f"{index}. "
            for j in i:
                s += f"{j} "
            f.write(f"{s}\n")
            index += 1

        f.close()

        f = open("reshumim.txt", "w", encoding='utf-8')
        index = 1
        for i in reshumim:
            f.write(f"{i[0]} {i[1]}\n")
            index += 1
        f.close()

        return reshumim

    def save_excel_file(self, array):
        a = []
        for j in range(len(array[0])):
            l = []

            for i in array:
                l.append(i[len(array[0]) - j - 1])
            a.append(l)

        df = pd.DataFrame(a).T
        df.to_excel(excel_writer="result.xlsx")

    def sort_file(self, path):
        f = open(path, "r", encoding='utf-8')
        names = f.readlines()
        f.close()

        names.sort()

        f = open(path, "w", encoding='utf-8')
        index = 1
        for i in names:
            f.write(i)
            index += 1
        f.close()


def main():
    t = TimeSaver("gdud.txt", "ז")
    t.save_my_time()


if __name__ == '__main__':
    main()
