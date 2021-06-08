import os
from tkinter import Frame, Tk, BOTH, Text, Menu, END, messagebox, Toplevel, Label, Scrollbar, RIGHT, Y, Button, \
    Radiobutton, StringVar, PhotoImage
from tkinter import filedialog

from xlutils.copy import copy

from fakturownia import get_last_12m_invoices
from functions import return_xls_sheet, return_xls_row, return_customized_xls_header, return_dict_from_lists, \
    return_split_dist, check_tax_numbers, return_invoice_no, return_invoice, compare_json_xls, return_date_time, \
    output_filename, return_xls_workbook


def get_compile_date():
    compilation_date = '(08-06-2021)'
    message = "Wersja 1.30 " + compilation_date

    messagebox.showinfo(title="Fakturownia",
                        message=message)


class InvoiceGui(Frame):

    def __init__(self):
        super().__init__()

        self.initUI()

    def refresh(self):
        self.destroy()
        self.__init__()

    def initUI(self):
        self.master.title("Fakturowania - DEMO")
        self.pack(fill=BOTH, expand=1)

        menubar = Menu(self.master)
        self.master.config(menu=menubar)

        fileMenu = Menu(menubar)
        fileMenu.add_command(label="Otwórz", command=self.on_open)
        fileMenu.add_command(label="Informacje", command=get_compile_date)
        menubar.add_cascade(label="Plik", menu=fileMenu)

        # vertical scroll bar
        v = Scrollbar(self)
        v.pack(side=RIGHT, fill=Y)

        self.txt = Text(self, yscrollcommand=v.set)
        self.txt.pack(fill=BOTH, expand=1)

        v.config(command=self.txt.yview)

    def console_output(self, text):
        text = text + '\n'
        self.txt.insert(END, text)

    def on_open(self):
        ftypes = [('Pliki Ms Excel', '*.xls'), ]
        dlg = filedialog.Open(self, filetypes=ftypes)
        fl = dlg.show()

        if fl != '':
            # create timestamp for the session
            timestamp = return_date_time()

            # correctly file path
            f_path = os.path.abspath(fl)

            self.console_output('Otwarto plik: ' + f_path)

            # messagebox.showinfo("Fakturowania", "Poprawnie wczytano plik")
            workbook = return_xls_workbook(f_path)
            sheet = return_xls_sheet(workbook)

            # get customized xls header
            xls_header = return_customized_xls_header(return_xls_row(sheet, 0))

            # initialize new xls workbook and sheet
            # copy rows between sheets
            output_wb = copy(workbook)
            output_ws = output_wb.get_sheet(0)

            # [1:] is needed for skipping first iteration
            # the first iteration has only column names
            for x in range(sheet.nrows)[1:]:
                # updating view
                self.update_idletasks()

                print('Row [' + str(x+1) + "]")
                self.console_output('Wiersz [' + str(x+1) + "]")

                # get row from xls file
                row = return_xls_row(sheet, x)

                # convert row to dictionary
                values_dict = return_dict_from_lists(xls_header, row)
                source_dict = return_split_dist(values_dict)
                print('Clean invoice')
                print(source_dict)

                # verifying is tax number is correctly get from xls file
                # sometimes tax number is blank or getting from a text file with an error (in source xls file)
                if check_tax_numbers(source_dict)['status'] == 'warning':
                    self.console_output('\t===Uwaga=== Napotkano błąd przy nr NIP w pliku"')
                    self.console_output('\t===Uwaga=== Zmieniono: ' + str(source_dict['NIP'])
                                        + ' na: ' + str(check_tax_numbers(source_dict)['new_val']))
                    source_dict['NIP'] = check_tax_numbers(source_dict)['new_val']

                elif check_tax_numbers(source_dict)['status'] == 'error':
                    # TODO do something if key not found
                    self.console_output('\t===Błąd=== Brak informacji o nr NIP lub płatność PAYPRO')
                    self.console_output('\t===Błąd=== ' + check_tax_numbers(source_dict)['message'])
                    output_ws.write(x, 7, 'error')
                    output_ws.write(x, 8, check_tax_numbers(source_dict)['message'])
                    continue

                # trying get invoice number
                invoice_number = return_invoice_no(source_dict)
                print('Invoice: ')
                print(invoice_number)

                if invoice_number['status'] == 'success':
                    print('Dictionary: ')
                    print(source_dict)
                    tax_id = source_dict['NIP']

                    json_output = get_last_12m_invoices(tax_id, f_path)

                    invoice = return_invoice(json_output['val'], invoice_number['val'])
                    print('Invoice number: ')
                    print(invoice)
                    self.console_output('\t===Sukces=== Poprawnie odczytano/zmieniono nr faktury"')

                    if invoice['status'] == 'success':
                        print('Payment: ')
                        print(compare_json_xls(invoice, source_dict))
                        self.console_output('\t===Sukces=== Wykonano odczytano/zmieniono status faktury"')
                        self.console_output('\t===Sukces=== ' + compare_json_xls(invoice, source_dict)['message'])
                        output_ws.write(x, 6, invoice['val']['number'][0:3])
                        output_ws.write(x, 7, invoice['status'])
                        output_ws.write(x, 8, compare_json_xls(invoice, source_dict)['message'])

                    elif invoice['status'] == 'error':
                        # TODO when we can not get invoice data
                        self.console_output('\t===Błąd=== Napotkano problem przy próbie uzyskania danych faktury"')
                        self.console_output('\t===Błąd=== ' + invoice['message'])
                        output_ws.write(x, 7, 'error')
                        output_ws.write(x, 8, invoice['message'])
                        continue

                elif invoice_number['status'] == 'error':
                    # TODO when we can not get properly invoice number
                    self.console_output('\t===Błąd=== Napotkano problem przy próbie uzyskania nr faktury"')
                    self.console_output('\t===Błąd=== ' + invoice_number['message'])
                    output_ws.write(x, 7, 'error')
                    output_ws.write(x, 8, invoice_number['message'])
                    continue

            output_xls_filename = output_filename(f_path, timestamp)

            try:
                output_wb.save(output_xls_filename)
                self.console_output('Plik wynikowy ===> ' + output_xls_filename)
            except Exception as e:
                self.console_output('\t===Błąd=== Problem z zapisem pliku wynikowego')
                self.console_output('\t===Błąd=== ' + str(e.__class__))

        else:

            messagebox.showerror("Fakturowania", "Błąd przy wczytwaniu pliku!")

        self.console_output('***Koniec pliku***')
        messagebox.showinfo("Fakturowania", "Skończono pracę z bieżącym plikiem")


def main():
    root = Tk()
    root.call()
    root.tk.call('wm', 'iconphoto', root.w, PhotoImage(file='settings/analytics.png'))
    ex = InvoiceGui()
    root.geometry("650x350+300+300")
    root.mainloop()


if __name__ == '__main__':
    main()
