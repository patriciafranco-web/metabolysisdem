import os
import sqlite3
import datetime
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = Flask(__name__)
app.secret_key = "clave_secreta_clinica_metabolica"
DB_NAME = "clinica_metabolica.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pacientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            edad INTEGER,
            genero TEXT,
            telefono TEXT,
            fecha_registro TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS antropometria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            paciente_id INTEGER,
            fecha TEXT,
            peso REAL,
            altura REAL,
            imc REAL,
            grasa_corporal REAL,
            masa_muscular REAL,
            cintura REAL,
            cadera REAL,
            FOREIGN KEY(paciente_id) REFERENCES pacientes(id)
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS clinica (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            paciente_id INTEGER,
            fecha TEXT,
            glucemia_ayunas REAL,
            hba1c REAL,
            colesterol_total REAL,
            hdl REAL,
            ldl REAL,
            trigliceridos REAL,
            presion_sistolica INTEGER,
            presion_diastolica INTEGER,
            FOREIGN KEY(paciente_id) REFERENCES pacientes(id)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

@app.route('/')
def index():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM pacientes ORDER BY id DESC")
    pacientes = cursor.fetchall()
    conn.close()
    return render_template('index.html', pacientes=pacientes)

@app.route('/agregar_paciente', methods=['POST'])
def agregar_paciente():
    nombre = request.form['nombre']
    edad = request.form['edad']
    genero = request.form['genero']
    telefono = request.form['telefono']
    fecha_reg = datetime.date.today().strftime('%Y-%m-%d')
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO pacientes (nombre, edad, genero, telefono, fecha_registro) VALUES (?, ?, ?, ?, ?)",
                   (nombre, edad, genero, telefono, fecha_reg))
    conn.commit()
    conn.close()
    flash("Paciente registrado exitosamente.", "success")
    return redirect(url_for('index'))

@app.route('/paciente/<int:p_id>')
def detalle_paciente(p_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM pacientes WHERE id = ?", (p_id,))
    paciente = cursor.fetchone()
    
    cursor.execute("SELECT * FROM antropometria WHERE paciente_id = ? ORDER BY fecha DESC", (p_id,))
    antro_list = cursor.fetchall()
    
    cursor.execute("SELECT * FROM clinica WHERE paciente_id = ? ORDER BY fecha DESC", (p_id,))
    clinica_list = cursor.fetchall()
    conn.close()
    
    if not paciente:
        flash("Paciente no encontrado.", "error")
        return redirect(url_for('index'))
        
    return render_template('detalle.html', paciente=paciente, antro_list=antro_list, clinica_list=clinica_list)

@app.route('/agregar_antropometria/<int:p_id>', methods=['POST'])
def agregar_antropometria(p_id):
    peso = float(request.form['peso'])
    altura = float(request.form['altura'])
    grasa = float(request.form['grasa'])
    musculo = float(request.form['musculo'])
    cintura = float(request.form['cintura'])
    cadera = float(request.form['cadera'])
    imc = peso / (altura ** 2)
    fecha = datetime.date.today().strftime('%Y-%m-%d')
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO antropometria (paciente_id, fecha, peso, altura, imc, grasa_corporal, masa_muscular, cintura, cadera)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (p_id, fecha, peso, altura, imc, grasa, musculo, cintura, cadera))
    conn.commit()
    conn.close()
    flash(f"Antropometría guardada. IMC calculado: {imc:.2f}", "success")
    return redirect(url_for('detalle_paciente', p_id=p_id))

@app.route('/agregar_clinica/<int:p_id>', methods=['POST'])
def agregar_clinica(p_id):
    glucemia = float(request.form['glucemia'])
    hba1c = float(request.form['hba1c'])
    col_total = float(request.form['col_total'])
    hdl = float(request.form['hdl'])
    ldl = float(request.form['ldl'])
    trig = float(request.form['trig'])
    pas = int(request.form['pas'])
    pad = int(request.form['pad'])
    fecha = datetime.date.today().strftime('%Y-%m-%d')
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO clinica (paciente_id, fecha, glucemia_ayunas, hba1c, colesterol_total, hdl, ldl, trigliceridos, presion_sistolica, presion_diastolica)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (p_id, fecha, glucemia, hba1c, col_total, hdl, ldl, trig, pas, pad))
    conn.commit()
    conn.close()
    flash("Registro clínico diabetológico guardado exitosamente.", "success")
    return redirect(url_for('detalle_paciente', p_id=p_id))

@app.route('/calculadoras', methods=['GET', 'POST'])
def calculadoras():
    homa_res = None
    gcb_res = None
    icr_res = None
    
    if request.method == 'POST':
        tipo = request.form.get('tipo')
        if tipo == 'homa':
            g = float(request.form['glucemia'])
            i = float(request.form['insulina'])
            homa = (g * i) / 405
            homa_res = f"{h:.2f} - {'Resistencia Elevada (>2.5)' if homa > 2.5 else 'Normal (<=2.5)'}"
        elif tipo == 'gcb':
            p = float(request.form['peso'])
            a = float(request.form['altura'])
            e = int(request.form['edad'])
            gen = request.form['genero']
            if gen == 'Masculino':
                gcb = 66.5 + (13.75 * p) + (5.003 * a) - (6.75 * e)
            else:
                gcb = 655.1 + (9.563 * p) + (1.85 * a) - (4.676 * e)
            gcb_res = f"{gcb:.0f} kcal/día"
        elif tipo == 'icr':
            cin = float(request.form['cintura'])
            alt = float(request.form['altura'])
            icr = cin / alt
            icr_res = f"{icr:.2f} - {'Riesgo aumentado (>0.5)' if icr > 0.5 else 'Riesgo bajo (<=0.5)'}"

    return render_template('calculadoras.html', homa_res=homa_res, gcb_res=gcb_res, icr_res=icr_res)

@app.route('/whatsapp')
def whatsapp_sim():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM pacientes")
    pacientes = cursor.fetchall()
    conn.close()
    return render_template('whatsapp.html', pacientes=pacientes)

@app.route('/generar_pdf/<int:p_id>')
def generar_pdf(p_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM pacientes WHERE id = ?", (p_id,))
    paciente = cursor.fetchone()
    conn.close()
    
    if not paciente:
        flash("Paciente no encontrado para generar PDF.", "error")
        return redirect(url_for('index'))
        
    filename = f"Informe_Clinico_{paciente[1].replace(' ', '_')}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#1f77b4'),
        spaceAfter=12
    )
    
    story.append(Paragraph("Informe Clínico Integral - Diabetes y Metabolismo", title_style))
    story.append(Paragraph(f"<b>Paciente:</b> {paciente[1]}", styles['Normal']))
    story.append(Paragraph(f"<b>Edad:</b> {paciente[2]} | <b>Género:</b> {paciente[3]}", styles['Normal']))
    story.append(Paragraph(f"<b>Fecha de Emisión:</b> {datetime.date.today()}", styles['Normal']))
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("<b>Recomendaciones Médicas y Plan de Tratamiento:</b>", styles['Heading2']))
    story.append(Paragraph("1. Control estricto de carbohidratos y glucemia capilar en ayunas.<br/>2. Actividad física aeróbica moderada de 30 minutos diarios.<br/>3. Hidratación adecuada y seguimiento con perfil lipídico mensual.", styles['Normal']))
    story.append(Spacer(1, 20))
    
    data = [
        ['Parámetro Clínico', 'Valor Objetivo', 'Estado'],
        ['HbA1c', '< 7.0%', 'En Seguimiento'],
        ['Presión Arterial', '< 130/80 mmHg', 'Controlado'],
        ['IMC', '18.5 - 24.9 kg/m²', 'Evaluado']
    ]
    t = Table(data, colWidths=[150, 150, 150])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#f2f2f2')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
    ]))
    story.append(t)
    doc.build(story)
    
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)