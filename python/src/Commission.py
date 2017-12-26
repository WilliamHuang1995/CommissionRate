from tkinter import *
from tkinter import filedialog
import openpyxl
import getpass


class GUI:
    def __init__(self, master):
        self.master = master
        master.title("Commission Calculator")

        self.select_button = Button(master, text="Select Input", command=self.select_input).grid(row=0)
        self.input_label = Label(master, text="", anchor="w", width=50)
        self.input_label.grid(row=0, column=1)
        self.output_button = Button(master, text="Select Output", command=self.select_output).grid(row=1)
        self.output_label = Label(master, text="", anchor="w", width=50)
        self.output_label.grid(row=1, column=1)
        self.run_button = Button(master, text="Run Program!", command=main).grid(row=2)
        self.close_button = Button(master, text="Close", command=master.quit).grid(row=2, column=1)

    def select_input(self):
        global path
        user = getpass.getuser()
        fn = filedialog.askopenfilename(initialdir="C:/Users/%s/Documents" % user, title="Select file",
                                        filetypes=(("Excel 2010 Files", "*.xlsx"), ("All Files", "*.*")))
        self.input_label['text'] = fn
        path = self.input_label['text']
        return fn

    def select_output(self):
        global dest
        fn = filedialog.askdirectory()
        self.output_label['text'] = fn + '/Output.xlsx'
        dest = self.output_label['text']
        return fn


def main():
    global my_gui
    if not path or not dest:
        print('input or output not set properly!')
        return
    # open workbook with only data in cells
    wb = openpyxl.load_workbook(path, data_only=True)
    # get first worksheet name
    wsn = wb.get_sheet_names()[0]
    # so I can get worksheet
    ws = wb.get_sheet_by_name(wsn)
    dictionary = process_data(ws)
    # create workbook
    owb = openpyxl.Workbook()
    output_data(dictionary, owb)
    # remove default sheet that Workbook creates
    owb.remove_sheet(owb.get_sheet_by_name('Sheet'))
    owb.save(filename=dest)
    # when successfully executed
    my_gui.output_label['text'] = 'Success'
    my_gui.input_label['text'] = 'Success'


def process_data(ws):
    employee_dict = {}
    # loop through row and columns
    for row in ws['A{}:J{}'.format(ws.min_row + 1, ws.max_row)]:
        employee = row[0].value.title()
        customer = row[2].value
        product = row[4].value
        revenue = row[9].value or 0
        # check if contain keywords/ put this in config file later
        if any(p in product for p in ('硬化劑', '發泡劑', '設備', '薄膜')):
            # check if employee already added
            if employee not in employee_dict.keys():
                employee_dict[employee] = {customer: {product: revenue}}
            else:
                # check if customer already added
                customers_dict = employee_dict[employee]
                # if there is no existing customer
                if customer not in customers_dict.keys():
                    customers_dict.update({customer: {product: revenue}})
                else:
                    products_dict = customers_dict[customer]
                    # if there is no existing product
                    if product not in products_dict.keys():
                        products_dict[product] = revenue
                    else:
                        products_dict[product] += revenue
    return employee_dict


def output_data(dictionary, owb):
    header = ['', '業務', '客戶', '產品', '金額', '2%佣金', '3%佣金', '5%佣金', 'Subtotal']
    # for each employee listed
    for employee in dictionary:
        ows = owb.create_sheet(title=employee)
        # insert header
        for col in range(1, len(header)):
            ows.cell(row=1, column=col).value = header[col]
        ows.cell(row=2, column=1).value = employee
        # row starts at 2
        # column starts at 2
        row = 2
        # iterate thru companies
        for company in dictionary[employee]:
            col = 3
            ows.cell(row=row, column=2).value = company
            for product in dictionary[employee][company]:
                revenue = dictionary[employee][company][product]
                values = [product, revenue, revenue * 0.02, revenue * 0.03, revenue * 0.05]
                for i in range(len(values)):
                    ows.cell(row=row, column=col + i).value = values[i]
                row += 1


if __name__ == '__main__':
    path = ''
    dest = ''
    root = Tk()
    my_gui = GUI(root)
    root.mainloop()
    root.destroy()
