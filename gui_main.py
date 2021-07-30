import os
from tkinter import Frame, Tk, BOTH, Text, Menu, END, messagebox, Toplevel, Label, Scrollbar, RIGHT, Y, Button, \
    Radiobutton, StringVar, PhotoImage, SUNKEN, W, X, BOTTOM, LEFT
from tkinter import filedialog

import requests
from xlutils.copy import copy

from fakturownia import get_last_12m_invoices, wait_for_connect, check_divider
from functions import return_xls_sheet, return_xls_row, return_customized_xls_header, return_dict_from_lists, \
    return_split_dist, check_tax_numbers, return_invoice_no, return_invoice, compare_json_xls, return_date_time, \
    output_filename, return_xls_workbook, return_invoice_by_status


# class for new dialog window, where we choose value from radiobutton
class MyDialog(object):
    def __init__(self, parent, current_value, raw_title, status, values):

        self.toplevel = Toplevel(parent)
        self.toplevel.minsize(600, 150)
        self.var = StringVar()
        label_text = "Lista faktur ze statusem: " + status
        status_label = Label(self.toplevel, text=label_text)
        status_label.pack(side="top", fill="x")

        if current_value is not None:
            label_text = current_value
            current_status_label = Label(self.toplevel, text=label_text, font='bold')
            current_status_label.pack(side="top", fill="x")

        label_text = raw_title
        raw_title_label = Label(self.toplevel, text=label_text, font='bold')
        raw_title_label.pack(side="top", fill="x")

        blank_label = Label(self.toplevel, text='')
        blank_label.pack(side="top", fill="x")

        # radiobutton list from dictionary
        for val in values:
            r = Radiobutton(
                self.toplevel,
                text=val[0],
                value=val[1],
                variable=self.var
            )
            r.pack(fill='x')

        blank_label = Label(self.toplevel, text='')
        blank_label.pack(side="top", fill="x")
        button = Button(self.toplevel, text="Wybierz", command=self.toplevel.destroy)
        button.pack()

    def show(self):
        self.toplevel.deiconify()
        self.toplevel.wait_window()
        value = self.var.get()
        return value


def get_compile_date():
    compilation_date = '(08-07-2021)'
    message = "Wersja 1.36PL " + compilation_date

    messagebox.showinfo(title="Fakturownia",
                        message=message)



class InvoiceGui(Frame):

    def status_bar(self):
        self.frame = Frame(padx=2, pady=2)
        self.frame.pack(side=BOTTOM)

        self.lblTitle = Label(self.frame, text="Status: ")
        self.lblTitle.pack(side=LEFT)

        self.lblStatusBar = Label(self.frame, text="Welcome", width=20, bd=1, relief=SUNKEN)
        self.lblStatusBar.pack(side=RIGHT)

    def status_bar_value(self, value):
        self.lblStatusBar.config(text=value)

    def __init__(self):
        super().__init__()
        self.status_bar()
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

                text = str(x) + '/' + str(sheet.nrows-1)
                self.status_bar_value(text)

                try:
                    # updating view
                    self.update_idletasks()

                    # print('Row [' + str(x+1) + "]")
                    self.console_output('Wiersz [' + str(x + 1) + "]")

                    # Wait for reconnect
                    if check_divider(int(x)):
                        text = str(x) + '/' + str(sheet.nrows - 1) + ' (Oczekiwanie...)'
                        self.status_bar_value(text)
                        self.update_idletasks()
                        wait_for_connect()

                    # get row from xls file
                    row = return_xls_row(sheet, x)

                    # convert row to dictionary
                    values_dict = return_dict_from_lists(xls_header, row)
                    source_dict = return_split_dist(values_dict)
                    # print('Clean invoice')
                    # print(source_dict)

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
                    # print('Invoice: ')
                    # print(invoice_number)

                    tax_id = source_dict['NIP']

                    json_output = get_last_12m_invoices(tax_id, f_path)

                    if invoice_number['status'] == 'success':
                        invoice = return_invoice(json_output['val'], invoice_number['val'])
                        # print('Invoice number: ')
                        # print(invoice)
                        self.console_output('\t===Sukces=== Poprawnie odczytano/zmieniono nr faktury"')

                        if invoice['status'] == 'success':
                            # print('Payment: ')
                            # print(compare_json_xls(invoice, source_dict))
                            self.console_output('\t===Sukces=== Wykonano odczytano/zmieniono status faktury"')
                            self.console_output('\t===Sukces=== ' + compare_json_xls(invoice, source_dict)['message'])
                            output_ws.write(x, 6, invoice['val']['number'][1:3])
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

                        # when we can not get properly invoice number we are trying find one unpaid invoice
                        self.console_output('\t===Błąd=== Napotkano problem przy próbie uzyskania nr faktury')
                        self.console_output('\t===Błąd=== ' + invoice_number['message'])
                        self.console_output('\t===UWAGA=== Podjęto próbę znalezienia nieopłaconej faktury')
                        issued_invoice = return_invoice_by_status(json_output['val'], 'issued')

                        if issued_invoice['status'] == 'success' and len(issued_invoice['val']) == 1:
                            self.console_output('\t===UWAGA=== Odnaleziono nieopłaconą fakturę')

                            print(json_output)

                            source_dict.update({'Tytuł': issued_invoice['val'][0]['number']})

                            invoice_number = return_invoice_no(source_dict)
                            print(invoice_number)

                            text = issued_invoice['val'][0]['number']
                            print(text)

                            invoice = return_invoice(json_output['val'], invoice_number['val'])
                            print(invoice)

                            self.console_output('\t===Sukces=== Wykonano odczytano/zmieniono status faktury"')
                            self.console_output('\t===Sukces=== ' + compare_json_xls(invoice, source_dict)['message'])
                            output_ws.write(x, 6, invoice['val']['number'][1:3])
                            output_ws.write(x, 7, invoice['status'])
                            output_ws.write(x, 8, compare_json_xls(invoice, source_dict)['message'])

                        else:
                            self.console_output('\t===UWAGA=== Napotkano na problem z odnaleziemiem nieopłaconej faktury '
                                                'lub znaleziono więcej niż jedną fakturę')
                            self.console_output('\t===Błąd=== ' + issued_invoice['message'])

                            output_ws.write(x, 7, 'error')
                            output_ws.write(x, 8, invoice_number['message'])
                            continue

                except Exception as e:
                    if str(e.__class__) == "<class 'requests.exceptions.ConnectionError'>":
                        self.console_output("")
                        self.console_output('>>>>>>>>>>Błąd<<<<<<<<<< ')
                        self.console_output('"Kod błędu: " ' + "<class 'requests.exceptions.ConnectionError'>")
                        output_ws.write(x, 7, 'fatal error')
                        output_ws.write(x, 8, "Wykryto problem z połączeniem internetowym!")
                        self.console_output("")
                        break
                    else:
                        self.console_output("")
                        self.console_output('>>>>>>>>>>Błąd<<<<<<<<<< ')
                        self.console_output('"Kod błędu: " ' + str(e.__class__))
                        self.console_output("")
                        output_ws.write(x, 7, 'error')
                        output_ws.write(x, 8, str(e.__class__))
                        continue

            output_xls_filename = output_filename(f_path, timestamp)

            try:
                output_wb.save(output_xls_filename)
                self.console_output("")
                self.console_output('Plik wynikowy ===> ' + output_xls_filename)
                self.console_output("")
            except Exception as e:
                self.console_output("")
                self.console_output('>>>>>>>>>>Błąd<<<<<<<<<< Problem z zapisem pliku wynikowego')
                self.console_output('"Kod błędu: " ' + str(e.__class__))
                self.console_output("")

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
