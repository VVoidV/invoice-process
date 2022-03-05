from email.policy import default
import pdfplumber
import openpyxl

ROOT = "D:\\Chinatelecom\\test"
XML_PATH = "D:\\Chinatelecom\\test\\test_2.xlsx"
PDF_PATH = "D:\\Chinatelecom\\test\\Q4"
SHEET_NAME = "工作表1"

values = [1000, 900, 800, 700, 600, 500, 400, 300, 200, 100, 50]
class Need:
        def __init__(self, amount, step):
                self.amount = amount
                self.detail = self.split_amount(amount, step)
        
        def split_amount(self, amount, step):
                remain = amount
                result = {}
                if amount == 0:
                        return result
                for value in values:
                        if value < step:
                                print("illegal need")
                                raise

                        result[value] = remain // value
                        remain -= result[value] * value

                        if remain == 0:
                                break

                assert(remain == 0) 
                return result
        def __repr__(self):
                return self.__str__()

        def __str__(self):
                res = " "
                for key, val in self.detail.items():
                        res += str(key) + ": " + str(val) +  " 张,"
                return res

class Person:
        def __init__(self, name, travel, cellphone):
                self.name = name or ""
                self.travel = travel or 0
                self.cellphone = cellphone or 0
                self.travel_need = Need(self.travel, 100)
        
        def __repr__(self):
                return self.__str__()

        def __str__(self):
                return self.name + ", 交通:" + str(self.travel) + ",通信: " + str(self.cellphone) + ",交通: " + str(self.travel_need)
class Invoice:
        def __init__(self, id, amount, path):
                self.id = id
                self.amoumt = amount
                self.path = path

        def __repr__(self):
                return self.__str__()

        def __str__(self):
                return str(self.id) + ", 金额:" + str(self.amount) + " 路径: " + str(self.path)


class Dispather:
        def __init__(self, root, xml_path, pdf_path):
                self.root = root
                self.xml_path = xml_path
                self.pdf_path = pdf_path
                self.employees = {}
                self.sheet_name = SHEET_NAME

        def load_xml(self):
                wb = openpyxl.load_workbook(self.xml_path, data_only=True)
                sheet = wb["工作表1"]

                # find colume
                for col in sheet.iter_cols():
                        if col[0].value == "姓名":
                                name_idx = col[0].col_idx
                        elif col[0].value == "交通":
                                travel_idx = col[0].col_idx
                        elif col[0].value == "通信":
                                phone_idx = col[0].col_idx
                        else:
                                continue

                for row in sheet.iter_rows(min_row=2):
                        name = row[name_idx - 1].value
                        if not name:
                                continue
                        travel = row[travel_idx - 1].value
                        phone = row[phone_idx - 1].value
                        p = Person(name, travel, phone)
                        self.employees[name] = p
                        print(p)
        
        def parse_pdf(self):
            with pdfplumber.open("D:\\ChinaTelecom\\test\\invoice_travel.pdf") as pdf:
                page = pdf.pages[0]
                pass

        
        def load_invoices(self):
                self.parse_pdf()
                pass

        def dispatch(self):
                pass

        def count_need(self):
                travel_need = dict(zip(values, [0]*len(values)))
                for entry in self.employees.values():
                        for key, value in entry.travel_need.detail.items():
                                travel_need[key] += value
                
                travel_total = 0
                for key, val in travel_need.items():
                        travel_total += key * val
                print("交通总额: ", travel_total)
                print(travel_need)
                






        


        
if __name__ == "__main__":
        worker = Dispather(ROOT, XML_PATH, PDF_PATH)                
        worker.load_xml()
        worker.count_need()
        pass
        # worker.load_invoices()

        # need = Need(4550, 50)
