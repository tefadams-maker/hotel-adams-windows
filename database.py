from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import hashlib

db = SQLAlchemy()

class Habitacion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    numero = db.Column(db.String(10), unique=True, nullable=False)
    tipo = db.Column(db.String(50), nullable=False)
    piso = db.Column(db.String(20), nullable=False)
    precio_base = db.Column(db.Float, nullable=False)
    estado_limpieza = db.Column(db.String(20), default='Limpia')
    estado_ocupacion = db.Column(db.String(20), default='Disponible')

class Huesped(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    apellido = db.Column(db.String(100))
    dni = db.Column(db.String(20), unique=True, nullable=False)
    celular = db.Column(db.String(20))
    email = db.Column(db.String(100))
    nacionalidad = db.Column(db.String(50), default='Perú')
    
class Reserva(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    habitacion_id = db.Column(db.Integer, db.ForeignKey('habitacion.id'), nullable=False)
    huesped_id = db.Column(db.Integer, db.ForeignKey('huesped.id'), nullable=False)
    fecha_entrada = db.Column(db.DateTime, default=datetime.now)
    fecha_salida = db.Column(db.DateTime)
    precio_total = db.Column(db.Float, nullable=False)
    precio_pagado = db.Column(db.Float, default=0)
    metodo_pago = db.Column(db.String(50))
    estado = db.Column(db.String(20), default='Activa')
    observaciones = db.Column(db.Text)
    
    habitacion = db.relationship('Habitacion', backref='reservas')
    huesped = db.relationship('Huesped', backref='reservas')

class Usuario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password = db.Column(db.String(120), nullable=False)
    nombre_completo = db.Column(db.String(100), nullable=False)
    rol = db.Column(db.String(50), default='Recepcion')
    activo = db.Column(db.Boolean, default=True)

    def __init__(self, username, email, password, nombre_completo, rol='Recepcion'):
        self.username = username
        self.email = email
        self.password = hashlib.sha256(password.encode()).hexdigest()
        self.nombre_completo = nombre_completo
        self.rol = rol
    
    @staticmethod
    def hash_password(password):
        return hashlib.sha256(password.encode()).hexdigest()
    
    def check_password(self, password):
        return self.hash_password(password) == self.password

class Producto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    precio = db.Column(db.Float, nullable=False)
    stock = db.Column(db.Integer, default=0)
    categoria = db.Column(db.String(50), default='General')

class CargoExtra(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    reserva_id = db.Column(db.Integer, db.ForeignKey('reserva.id'), nullable=False)
    producto_id = db.Column(db.Integer, db.ForeignKey('producto.id'), nullable=False)
    cantidad = db.Column(db.Integer, nullable=False)
    precio_unitario = db.Column(db.Float, nullable=False)
    fecha = db.Column(db.DateTime, default=datetime.now)
    
    reserva = db.relationship('Reserva', backref=db.backref('cargos_extras', lazy=True))
    producto = db.relationship('Producto', backref=db.backref('ventas', lazy=True))