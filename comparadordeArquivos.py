import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xlsxwriter
from xlsxwriter import Workbook

listfile = []


class comparadorArquivos(tk.Tk):
    def __init__(self, parent):
        tk.Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()

    def initialize(self):

        lb1 = tk.Label(self, text='Arquivo 1', background='#dde', foreground='#000')
        lb1.place(x=10, y=20, width='150', height='20')

        bt1 = tk.Button(self, text='Inserir', foreground='#000', command=self.open)
        bt1.place(x=150, y=20, width='100', height='20')

        lb2 = tk.Label(self, text='Arquivo 2', background='#dde', foreground='#000')
        lb2.place(x=10, y=50, width='150', height='20')

        bt2 = tk.Button(self, text='Inserir', foreground='#000', command=self.open)
        bt2.place(x=150, y=50, width='100', height='20')

        bt3 = tk.Button(self, text='Comparar', foreground='#000', command=self.compare)
        bt3.place(x=80, y=100, width='150', height='20')

    def open(self):
        x = filedialog.askopenfilename(title='Por favor selecione o arquivo:')
        listfile.append(x)
        mensagem = "Arquivo " + str(len(listfile)) + " selecionado"
        messagebox.showinfo(message=mensagem)

    def compare(self):

        def RemoveSpace(tb, serie):
            try:
                return tb[serie].str.strip()
            except:
                return tb[serie]

        def rename(tb, change):
            for col in tb.columns:
                if col.find("wire") == -1:
                    novo_nome = col + "_" + change
                    tb = tb.rename({col: novo_nome}, axis=1)
            return (tb)

        arquivo1 = pd.read_csv(listfile[0])
        arquivo2 = pd.read_csv(listfile[1])

        if arquivo1.equals(arquivo2):
            messagebox.showinfo(message="os arquivos são iguais!")
        else:
            arquivo1.columns = arquivo1.columns.str.lower().str.strip()
            arquivo2.columns = arquivo2.columns.str.lower().str.strip()

            arquivo1 = arquivo1.rename(
                columns={'ref. 1': 'ref1', 'pos. 1': 'pos1', 'term 1': 'term1', 'seal 1': 'seal1', 'ref. 2': 'ref2',
                         'pos. 2': 'pos2', 'term 2': 'term2', 'seal 2': 'seal2'})
            arquivo2 = arquivo2.rename(
                columns={'ref. 1': 'ref1', 'pos. 1': 'pos1', 'term 1': 'term1', 'seal 1': 'seal1', 'ref. 2': 'ref2',
                         'pos. 2': 'pos2', 'term 2': 'term2', 'seal 2': 'seal2'})
            for col in arquivo1.columns:
                arquivo1[col] = RemoveSpace(arquivo1, col)

            for col in arquivo2.columns:
                arquivo2[col] = RemoveSpace(arquivo2, col)

            arquivo1 = rename(arquivo1, "old")
            arquivo2 = rename(arquivo2, "new")

            try:
                merge1 = arquivo2.merge(arquivo1, how='left')
                inserido = merge1[merge1[
                    ['cable_old', 'area_old', 'part_no_old', 'color_old', 'type_old', 'ref1_old', 'pos1_old',
                     'term1_old', 'seal1_old', 'ref2_old', 'pos2_old', 'term2_old', 'seal2_old', 'variant_old',
                     'assembly_old']].isna().all(1)]
                inserido = inserido.fillna('-')
                inserido['Status'] = 'ADDED'
            except:
                messagebox.showinfo(message='Não há itens em comum')

            merge2 = arquivo1.merge(arquivo2, how='left')
            excluido = merge2[merge2[
                ['cable_new', 'area_new', 'part_no_new', 'color_new', 'type_new', 'ref1_new', 'pos1_new', 'term1_new',
                 'seal1_new', 'ref2_new', 'pos2_new', 'term2_new', 'seal2_new', 'variant_new',
                 'assembly_new']].isna().all(1)]
            excluido = excluido.fillna('-')
            excluido['Status'] = 'REMOVED'

            inner = arquivo1.merge(arquivo2, how='inner')
            inner.loc[(inner.cable_old != inner.cable_new) | (inner.area_old != inner.area_new) |
                      (inner.part_no_old != inner.part_no_new) | (inner.color_old != inner.color_new) |
                      (inner.type_old != inner.type_new) | (inner.ref1_old != inner.ref1_new) |
                      (inner.pos1_old != inner.pos1_new) | (inner.term1_old != inner.term1_new) |
                      (inner.seal1_old != inner.seal1_new) | (inner.ref2_old != inner.ref2_new) |
                      (inner.pos2_old != inner.pos2_new) | (inner.term2_old != inner.term2_new) |
                      (inner.seal2_old != inner.seal2_new) | (inner.variant_old != inner.variant_new) |
                      (inner.assembly_old != inner.assembly_new), 'Status'] = 'MODIFIED'

            filtro = inner['Status'] == 'MODIFIED'
            modificado = inner[filtro]

            tabelaFinal = pd.concat([inserido, modificado, excluido])
            tabelaFinal = tabelaFinal.reset_index(drop=True)
            tabelaFinal = tabelaFinal[
                ['cable_old', 'cable_new', 'wire', 'area_old', 'area_new', 'part_no_old', 'part_no_new',
                 'color_old', 'color_new', 'type_old', 'type_new', 'ref1_old', 'ref1_new', 'pos1_old', 'pos1_new',
                 'term1_old', 'term1_new', 'seal1_old', 'seal1_new', 'ref2_old', 'ref2_new', 'pos2_old', 'pos2_new',
                 'term2_old', 'term2_new', 'seal2_old', 'seal2_new', 'variant_old', 'variant_new', 'assembly_old',
                 'assembly_new', 'Status']]

            writer = pd.ExcelWriter('ComparadorWireList.xlsx', engine='xlsxwriter')
            tabelaFinal.to_excel(writer, sheet_name='ComparadorWireList', index=False)

            workbook = writer.book
            worksheet = writer.sheets['ComparadorWireList']

            worksheet.set_column("A:AF", 12)

            rows = len(tabelaFinal.index)

            colorRange = "AF2:AF{}".format(rows + 1)

            green_format = workbook.add_format({'bg_color': '#C6EFCE',
                                                'font_color': '#006100'})

            yellow_format = workbook.add_format({'bg_color': '#FFEB9C',
                                                 'font_color': '#9C6500'})

            red_format = workbook.add_format({'bg_color': '#FFC7CE',
                                              'font_color': '#9C0006'})

            worksheet.conditional_format(colorRange, {'type': 'cell',
                                                      'criteria': 'equal to',
                                                      'value': '"ADDED"',
                                                      'format': green_format})

            worksheet.conditional_format(colorRange, {'type': 'cell',
                                                      'criteria': 'equal to',
                                                      'value': '"MODIFIED"',
                                                      'format': yellow_format})

            worksheet.conditional_format(colorRange, {'type': 'cell',
                                                      'criteria': 'equal to',
                                                      'value': '"REMOVED"',
                                                      'format': red_format})

            writer.save()

            messagebox.showinfo(message="Arquivo gerado com sucesso!")


if __name__ == "__main__":
    app = comparadorArquivos(None)
    app.title('Comparador de arquivos Wire List')
    app.geometry('300x150')
    app.configure(background='#dde')
    app.mainloop()



