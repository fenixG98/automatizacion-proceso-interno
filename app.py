from flask import Flask, request, render_template, url_for, flash
from werkzeug.utils import redirect, secure_filename, send_from_directory
import os
import sys
import xlrd
import xlsxwriter
import mysql.connector

UPLOAD_FOLDER = "./files/"
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = "SECRET_KEY"


# Classe que crea un fitxer d'importació
class Fitxer:
    def __init__(self, directori):
        self._directori = directori

    def get_dir(self):
        return self._directori


# Classe que crea un llistat de clients
class Clients:
    def __init__(self):
        self._connexio = self.set_con()
        self._clients_dic = self.set_dic()

    def set_con(self):
        con = 0

        try:
            con= mysql.connector.connect(auth_plugin='mysql_native_password', host="localhost", database="clients_E", user="admin", passwd="password", port=3306)
        except Exception as e:
            app.logger.error("No és possible connectar amb la base de dades.")
            print("Sortint del programa. No és possible connectar amb la BBDD. Error: " + str(e))

        return con

    def set_dic(self):
        print("Carregant clients al programa.")
        cursor = self.get_con().cursor()
        sql = 'SELECT client,E FROM clients_E'
        cursor.execute(sql)
        dades = cursor.fetchall()
        cursor.close()
        return {clau: valor for clau, valor in dades}

    def get_con(self):
        return self._connexio

    def get_dic(self):
        return self._clients_dic

    def get_e(self, client):
        e = 'ERROR'
        clients = self.get_dic()
        try:
            e = clients[client]
            print(client + ' ' + clients[client])
        except Exception as err:
            app.logger.error("No és possible trobar el client " + client + ". Error: " + str(err))
        return e

    def editar_e(self, fitxer, client, e):
        con = self.get_con()
        cursor = con.cursor()

        if (fitxer!=app.config['UPLOAD_FOLDER']):

            Import = xlrd.open_workbook(fitxer)
            ImportDades = Import.sheet_by_index(0)
            for x in range(ImportDades.nrows):
                sql = 'SELECT * FROM clients_E where client="' + ImportDades.cell_value(x, 0) + '";'
                cursor.execute(sql)
                dades = cursor.fetchall()
                if (dades == []):
                    sql = 'INSERT clients_E (client, E) VALUES ("' + ImportDades.cell_value(x,
                                                                                            0) + '","' + ImportDades.cell_value(
                        x, 1) + '")'
                    cursor.execute(sql)
                else:
                    sql = 'UPDATE clients_E SET E="' + ImportDades.cell_value(x,
                                                                              1) + '" WHERE client ="' + ImportDades.cell_value(
                        x, 0) + '";'
                    cursor.execute(sql)
                print(ImportDades.cell_value(x, 0) + ' ' + ImportDades.cell_value(x, 1))
        else:
            sql = 'SELECT * FROM clients_E where client="' + client + '";'
            cursor.execute(sql)
            dades = cursor.fetchall()

            if (dades == []):
                sql = 'INSERT clients_E (client, E) VALUES ("' + client + '","' + e + '")'
                cursor.execute(sql)
            else:
                sql = 'UPDATE clients_E SET E="' + e + '" WHERE client ="' + client + '";'
                cursor.execute(sql)
            print(client + ' ' + e)

        con.commit()
        cursor.close()

# Classe que gestiona Excels
class GestorExcels:

    def escriure_monitoritzacio(self, x, data, e, operador, temps):
        global ExportDades
        ExportDades.write(x, 0, data)
        ExportDades.write(x, 1, e)
        ExportDades.write(x, 2, operador)
        ExportDades.write(x, 3, temps)
        ExportDades.write(x, 4, 'Monitorització')

    def script_recursiu(self, dataAnt, lectura, escriptura, monitoritzacio):
        global clients

        if (lectura == ImportDades.nrows):
            self.escriure_monitoritzacio(escriptura, ImportDades.cell_value(lectura-1, 0), clients.get_e('[...]'), ImportDades.cell_value(1, 2), monitoritzacio)
            return

        dataAct = ImportDades.cell_value(lectura, 0)
        if (dataAct != dataAnt):
            self.escriure_monitoritzacio(escriptura, dataAnt, clients.get_e('[...]'), ImportDades.cell_value(1, 2), monitoritzacio)
            dataAnt = dataAct            
            return self.script_recursiu(dataAnt, lectura, escriptura+1, 480)

        ExportDades.write(escriptura, 0, ImportDades.cell_value(lectura, 0))                                                            # Data
        ExportDades.write(escriptura, 1, clients.get_e(ImportDades.cell_value(lectura, 1)))                                             # E
        ExportDades.write(escriptura, 2, ImportDades.cell_value(lectura, 2))                                                            # Operador
        ExportDades.write(escriptura, 3, ImportDades.cell_value(lectura, 5))                                                            # Temps
        ExportDades.write(escriptura, 4, str(ImportDades.cell_value(lectura, 3)) + ' - ' + str(ImportDades.cell_value(lectura, 4)))     # Tiquet
       
        monitoritzacio -= ImportDades.cell_value(lectura, 5)  

        return self.script_recursiu(dataAnt, lectura+1, escriptura+1, monitoritzacio)

    def script_principal(self, directoriExport):
        global fitxer
        global entorn
        global ImportDades

        ImportV = fitxer.get_dir()
        Import = xlrd.open_workbook(ImportV)
        ImportDades = Import.sheet_by_name('tiquets')

        for x in ('\/:*?"<>| '):
            for y in directoriExport:
                if (x == y):
                    app.logger.error("El nom del fitxer conté caràcters no admesos.")
                    return
        ExportV = './files/' + directoriExport + ".xlsx"
        Export = xlsxwriter.Workbook(ExportV)
        global ExportDades
        ExportDades = Export.add_worksheet()
        print("Llegint Excel")
        self.script_recursiu(ImportDades.cell_value(1, 0), 1, 0, 480)
        print("Escrivint Excel")
        Export.close()
        app.logger.info("PROCESSAMENT FINALITZAT")
        app.logger.info("Informe generat amb el nom: " + directoriExport + '.xlsx')


# Funcio principal del programa
def script(directoriExport, directoriImport):
    global clients
    global fitxer
    fitxer = Fitxer(directoriImport)
    clients = Clients()
    gestor = GestorExcels()
    gestor.script_principal(directoriExport)


# Funcio que valida el nom del fitxer
def validacio_nom_fitxer(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# http://localhost:5000/
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            app.logger.error("No has seleccionat cap fitxer d'importació")
            return redirect(request.url)
        file = request.files['file']

        if file.filename == '':
            app.logger.error("No has seleccionat cap fitxer d'importació")
            return redirect(request.url)
        if file and validacio_nom_fitxer(file.filename):
            global export
            export = request.form['nomInforme']
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            script(export, os.path.join(app.config['UPLOAD_FOLDER'], filename))
            uploads = os.path.join(app.config['UPLOAD_FOLDER'])
            export = export+'.xlsx'
            return send_from_directory(directory=uploads, path=export, environ=request.environ, as_attachment=True)

    return render_template('index.html')


# http://localhost:5000/easteregg
@app.route("/easteregg")
def easter_egg():
    return redirect(url_for('index'))


# http://localhost:5000/editar
@app.route('/editar', methods=['GET', 'POST'])
def editar():
    if request.method == 'POST':
        file = request.files['file']
        if 'file' not in request.files or file.filename == '':
            app.logger.error("No has seleccionat cap fitxer d'importació")
            print('no fitxer')
            filename = ''

        else:
            if file and validacio_nom_fitxer(file.filename):
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            else:
                filename = ''
                app.logger.error("Nom no admès")

            print('fitxer')

        clients = Clients()
        clients.editar_e(os.path.join(app.config['UPLOAD_FOLDER'], filename), request.form['client'], request.form['E'])
        return redirect(url_for('clients'))

    return render_template('editar.html')


# http://localhost:5000/clients
@app.route('/clients', methods=['GET', 'POST'])
def clients():
    clients = Clients()
    clients = clients.get_dic()
    print (clients)
    return render_template('clients.html', clients=clients)


# http://localhost:5000/asdfg
@app.errorhandler(404)
def pagina_no_trobada(error):
    return render_template('error_404.html', error=error)


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, use_reloader=False)