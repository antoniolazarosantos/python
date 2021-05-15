import pyodbc


class ConectarDB:
    def __init__(self):
        # Criando conexão.
        self.con = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                                  r'DBQ=C:\Mercado Livre\Áreas\Waves\Passagem de Turno\BDPassagemTurnoOutbound.accdb;')
        # Criando cursor.
        self.cur = self.con.cursor()

    def exibir_tabelas(self):
        for tabela in self.cur.tables(tableType='TABLE'):
            print(tabela.table_name)

    def consultar_registros(self):
        return self.cur.execute('SELECT * FROM Calendario order by 1').fetchall()


    def inserir_script_diario(self, pLista):
        # Calendário
        # self.cur.execute("select max(código) from calendario")
        # row = self.cur.fetchone()
        # id = 0
        # while row:
        #    id = row[0]
        #    row = self.cur.fetchone()

        print("Limpando dados da data...")
        lista = ["calendario", "sla", "produtividade", "[Unidades em Buffer]",
                 "[Unidades Processadas]", "waves"]
        lData = pLista[0].split("/")
        xData = lData[1]+'/'+lData[0]+'/'+lData[2]
        for t in lista:
            self.cur.execute("delete from {0} where data = #{1}#".format(t, xData))
        self.cur.commit()

        print("Iniciando calendário...")
        self.cur.execute("insert into calendario (código,data,target) "
                         "select max(código)+1,'{0}',{1} from calendario".format(pLista[0],pLista[1]))
        # SLA
        print("Iniciando SLA...")
        self.cur.execute("insert into sla (id,data) "
                         "select max(id)+1,'{0}' from sla".format(pLista[0]))
        # Produtividade
        print("Iniciando produtividade...")
        self.cur.execute("insert into produtividade (id,data) "
                         "select max(id)+1,'{0}' from produtividade".format(pLista[0]))
        # Unidades em Buffer
        print("Iniciando Unidades em Buffer...")
        lista = ["Planning", "Picking", "Packing", "Put to Wall", "Shipping"]
        for processo in lista:
            self.cur.execute("insert into [Unidades em Buffer] (processo,data,turno,buffer) "
                             "values ('{0}','{1}','Noturno',0)".format(processo, pLista[0]))
        # Unidades Processadas
        print("Iniciando Unidades Processadas...")
        lista = ["Picking", "Wall", "Packing", "Shipping"]
        for processo in lista:
            self.cur.execute("insert into [Unidades Processadas] (processo,data,turno,realizado) "
                             "values ('{0}','{1}','Noturno',0)".format(processo, pLista[0]))

        # Waves
        print("Iniciando Waves...")
        listaHora = ["22:00", "23:00", "00:00", "01:00", "02:00", "03:00", "04:00"]
        for hora in listaHora:
            self.cur.execute("insert into waves (processo,data,turno,intervalo,"
                             "mono,multi,emergência,retiro) "
                             "values ('Wave','{0}','Noturno','{1}',1,1,0,0)".format(pLista[0], hora))

        self.cur.commit()
        print("Processo finalizado com sucesso.")


if __name__ == '__main__':
    banco = ConectarDB()
    # banco.exibir_tabelas()
    data = str(input("Data de Processamento: "))
    target = int(input("Target: "))
    banco.inserir_script_diario([data,target])
    # print(banco.consultar_registros())
