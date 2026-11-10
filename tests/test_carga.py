import contextlib
import datetime
import io
from pathlib import Path
import tempfile
import unittest

import openpyxl
from preparar_carga import process


class CargaTest(unittest.TestCase):
    def fixture(self, path):
        w = openpyxl.Workbook()
        s = w.active
        s.title = 'Hoja1'
        s.append([])
        s.append([None, 'C.C', 'Sucursal', 'N° SUC.', 'N° Cliente', None, None, None, None, None, None,
                  'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre'])
        def source(client, code, values):
            s.append([None, None, 'Sucursal', code, client, None, None, None, None, None, None] + values)
        source('123-4', 1, [8, 0, 'PAGA FUNCIONARIO'])
        source('123-4', 1, [None, 0])
        source('GESTION INFRAESTRUCTURA', 2, [5])
        source('GASTO COMUN', 3, [7])
        source('999/456-7', 4, [9])
        source('777', 5, [10])
        source('DESCONOCIDO', 6, [3])
        source('222-2', 7, [1])
        source('222-2', 7, [2])
        c = w.create_sheet('BBDD N° Clientes')
        c.append(['Numero de cliente', 'Proveedor', 'Codigo', None, None, None, None, None, None, 'Sucursal'])
        for client, code in [('123-4',1),('GESTIONINFRA01',2),('GASTOCOMUN01',3),('456-7',4),('777-8',5),('222-2',7)]:
            c.append([client, 'Proveedor', code, None, None, None, None, None, None, 'Sucursal'])
        h = w.create_sheet('Base de datos Agua Reprocesada')
        h.append([None]*9 + ['Número de Cliente','Fecha','Número de documento','Consumo [m3]',
                            'Limite de consumo [m3]','Sobre consumo','Total a Pagar','Total Mes [$]','Saldo Anterior [$]','Mes','Año'])
        h.append([None]*9 + ['123-4',datetime.datetime(2026,1,1),None,99])
        w.save(path)
        w.close()

    def test_missing_periods_special_clients_zero_text_conflicts(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, dest = Path(tmp)/'in.xlsx', Path(tmp)/'out.xlsx'
            self.fixture(source)
            original = source.read_bytes()
            with contextlib.redirect_stdout(io.StringIO()):
                counts = process(source, dest)
            self.assertEqual(source.read_bytes(), original)
            self.assertEqual(counts['registros_para_cargar'], 6)
            self.assertEqual(counts['registros_ya_existentes'], 1)
            self.assertEqual(counts['registros_pendientes'], 3)
            self.assertEqual(counts['duplicados_internos_omitidos'], 1)
            self.assertEqual(counts['ceros_para_cargar'], 1)
            self.assertEqual(counts['textos_para_cargar'], 1)
            result = openpyxl.load_workbook(dest, read_only=True)
            rows = list(result['Carga J-R'].iter_rows(min_row=2, values_only=True))
            self.assertEqual({r[0] for r in rows}, {'123-4','GESTIONINFRA01','GASTOCOMUN01','456-7','777-8'})
            self.assertTrue(all(r[1].day == 28 for r in rows))
            self.assertTrue(all(r[5] is None for r in rows))
            result.close()

    def test_reject_incompatible_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, dest = Path(tmp)/'in.xlsx', Path(tmp)/'out.xlsx'
            w = openpyxl.Workbook()
            w.save(source)
            w.close()
            with self.assertRaisesRegex(ValueError, 'Faltan hojas'):
                process(source, dest)
            self.assertFalse(dest.exists())

    def test_selected_year(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, dest = Path(tmp)/'in.xlsx', Path(tmp)/'out.xlsx'
            self.fixture(source)
            with contextlib.redirect_stdout(io.StringIO()):
                counts = process(source, dest, 2027)
            self.assertEqual(counts['registros_ya_existentes'], 0)
            result = openpyxl.load_workbook(dest, read_only=True)
            self.assertTrue(all(r[1].year == 2027 for r in result['Carga J-R'].iter_rows(min_row=2,values_only=True)))
            result.close()
