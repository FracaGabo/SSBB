import openpyxl, re, unicodedata, collections, datetime, json
from pathlib import Path
def norm(v):
 if v is None: return ''
 if isinstance(v,float) and v.is_integer(): v=int(v)
 return re.sub(r'[^A-Z0-9]','',unicodedata.normalize('NFKD',str(v)).encode('ascii','ignore').decode().upper())
def present(v): return v is not None and (not isinstance(v,str) or bool(v.strip()))

def process(source, destination, year=2026):
 w=openpyxl.load_workbook(source,read_only=True,data_only=True)
 required = {'Hoja1', 'BBDD N° Clientes', 'Base de datos Agua Reprocesada'}
 missing = required - set(w.sheetnames)
 if missing:
  raise ValueError('Faltan hojas: ' + ', '.join(sorted(missing)))
 checks = [('Hoja1', 2, {4: 'N° Cliente', **{11+i: name for i,name in enumerate(['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre'])}}), ('BBDD N° Clientes', 1, {0: 'Numero de cliente', 2: 'Codigo'}), ('Base de datos Agua Reprocesada', 1, {9: 'Número de Cliente',10: 'Fecha',12: 'Consumo [m3]'})]
 for sn, rn, expected in checks:
  header = next(w[sn].iter_rows(min_row=rn,max_row=rn,values_only=True))
  for col, label in expected.items():
   if len(header)<=col or norm(header[col])!=norm(label):
    raise ValueError(f'Encabezado inesperado en {sn}, columna {col+1}: se esperaba {label}')
 cat=[r for r in w['BBDD N° Clientes'].iter_rows(min_row=2,values_only=True) if present(r[0])]
 byid=collections.defaultdict(list)
 for r in cat: byid[norm(r[0])].append(r)
 def match(r):
  ident=norm(r[4]); code=norm(r[3]); name=norm(r[2])
  candidates=byid.get(ident,[]) if ident not in ('','0') else []
  method='Número de cliente normalizado'
  if not candidates and '/' in str(r[4]):
   parts=[norm(x) for x in str(r[4]).split('/')]
   candidates=[c for c in cat if norm(c[0]) in parts and norm(c[2])==code]
   method='Componente del número compuesto + código sucursal'
  if not candidates and ident.isdigit():
   candidates=[c for c in cat if norm(c[2])==code and str(c[0]).rsplit('-',1)[0].strip()==str(r[4]).strip() and '-' in str(c[0])]
   method='Número sin dígito verificador + código sucursal'
  if not candidates and ('GESTION' in ident or 'GABRIEL' in ident or ident in ('GASTOCOMUN','REMARCADOR') or (ident==code and code not in ('','0'))):
   candidates=[c for c in cat if code not in ('','0') and norm(c[2])==code]
   method='Identificador especial: código de sucursal'
   family='GESTIONINFRA' if 'INFRA' in ident else 'GESTIONCOMUN' if 'COMUN' in ident and 'GESTION' in ident else 'GESTIONINT' if 'INTERNA' in ident else None
   if family: candidates=[c for c in candidates if norm(c[0]).startswith(family)]
  if len({norm(c[0]) for c in candidates})>1:
   narrowed=[c for c in candidates if code not in ('','0') and norm(c[2])==code]
   if narrowed: candidates=narrowed
  if len({norm(c[0]) for c in candidates})>1:
   narrowed=[c for c in candidates if norm(c[9])==name]
   if narrowed: candidates=narrowed
  ids={norm(c[0]) for c in candidates}
  return (candidates[0],method) if len(ids)==1 else (None,'Sin coincidencia' if not candidates else 'Correspondencia ambigua')
 existing=collections.defaultdict(list)
 for n,r in enumerate(w['Base de datos Agua Reprocesada'].iter_rows(min_row=2,values_only=True),2):
  if not present(r[9]): continue
  d=r[10]
  if isinstance(d,(datetime.datetime,datetime.date)): y,m=d.year,d.month
  elif isinstance(r[18],(int,float)) and isinstance(r[19],(int,float)): y,m=int(r[19]),int(r[18])
  else: continue
  existing[(norm(r[9]),y,m)].append((n,r[12]))
 proposed=collections.defaultdict(list); review=[]; excluded=[]; mapping=[]; all_records=[]; counts=collections.Counter()
 for n,r in enumerate(w['Hoja1'].iter_rows(min_row=3,values_only=True),3):
  vals=[(m,v) for m,v in enumerate(r[11:21],1) if present(v)]
  counts['celdas_con_datos']+=len(vals)
  if not vals: continue
  c,method=match(r)
  for m,value in vals:
   all_records.append(([c[0] if c is not None else r[4],datetime.datetime(year,m,28),None,value,None,None,None,None,None],method if c is None else None))
  if c is None:
   for m,v in vals: review.append([n,r[2],r[3],r[4],None,datetime.datetime(year,m,28),v,method])
   continue
  mapping.append([n,r[2],r[3],r[4],c[0],c[2],c[9],method])
  for m,v in vals:
   key=(norm(c[0]),year,m)
   if key in existing:
    excluded.append([n,r[2],c[0],datetime.datetime(year,m,28),v,', '.join(str(x[0]) for x in existing[key]),' | '.join(str(x[1]) for x in existing[key])])
   else: proposed[key].append((n,r,c,v,m,method))
 rows=[]; audit=[]
 for key,items in proposed.items():
  first=items[0]
  # Do not choose between different source values for the same client/month.
  if len({(type(x[3]).__name__,str(x[3])) for x in items})>1:
   for n,r,c,v,m,method in items: review.append([n,r[2],r[3],r[4],c[0],datetime.datetime(year,m,28),v,'Valores de origen distintos para el mismo cliente y mes'])
   continue
  n,r,c,v,m,method=first
  rows.append([c[0],datetime.datetime(year,m,28),None,v,None,None,None,None,None])
  audit.append([c[0],datetime.datetime(year,m,28),v,', '.join(str(x[0]) for x in items),r[2],method])
  counts['duplicados_internos_omitidos']+=len(items)-1
 rows.sort(key=lambda x:(norm(x[0]),x[1])); audit.sort(key=lambda x:(norm(x[0]),x[1]))
 headers=list(next(w['Base de datos Agua Reprocesada'].values))[9:18]
 out=openpyxl.Workbook(); out.remove(out.active)
 from openpyxl.styles import Font, PatternFill
 from openpyxl.utils import get_column_letter
 def sheet(name,head,data):
  s=out.create_sheet(name); s.append(head)
  for row in data:
   s.append(row)
   for cell in s[s.max_row]:
    if isinstance(cell.value,str): cell.data_type='s'
    if isinstance(cell.value,datetime.datetime): cell.number_format='dd/mm/yyyy'
  for cell in s[1]: cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='245A81')
  s.freeze_panes='A2'; s.auto_filter.ref=s.dimensions
  for i in range(1,len(head)+1):
   width=min(65,max(16,max(len(str(s.cell(j,i).value or '')) for j in range(1,min(s.max_row,150)+1))+2))
   s.column_dimensions[get_column_letter(i)].width=width
  return s
 sheet('Carga J-R',headers,rows)
 all_records.sort(key=lambda item:(norm(item[0][0]),item[0][1]))
 all_sheet=sheet('Todos J-R',headers,[record for record,reason in all_records])
 from openpyxl.comments import Comment
 for row_number,(record,reason) in enumerate(all_records,2):
  if reason:
   for cell in all_sheet[row_number]:
    cell.fill=PatternFill('solid',fgColor='FFF2CC')
   all_sheet.cell(row_number,1).comment=Comment('Cliente original de Hoja1 sin correspondencia segura: '+reason+'. Revisar Pendientes; no usar esta fila para cargar.', 'SSBB')
 sheet('Pendientes',['Fila Hoja1','Sucursal','Código sucursal','Cliente origen','Cliente catálogo','Fecha','Consumo o texto','Motivo'],review)
 sheet('Ya existentes',['Fila Hoja1','Sucursal','Cliente catálogo','Fecha propuesta','Valor Hoja1','Filas base Agua','Consumos base Agua'],excluded)
 sheet('Trazabilidad',['Cliente catálogo','Fecha','Consumo o texto','Filas Hoja1','Sucursal','Método de correspondencia'],audit)
 sheet('Correspondencias',['Fila Hoja1','Sucursal origen','Código origen','Cliente origen','Cliente catálogo','Código catálogo','Sucursal catálogo','Método'],mapping)
 counts.update({'registros_totales':len(all_records),'registros_para_cargar':len(rows),'registros_ya_existentes':len(excluded),'registros_pendientes':len(review),'ceros_para_cargar':sum(isinstance(r[3],(int,float)) and r[3]==0 for r in rows),'textos_para_cargar':sum(isinstance(r[3],str) for r in rows)})
 notes=[['Año',year],['Todos J-R','Todos los datos no vacíos de Hoja1, incluidos períodos existentes y repeticiones de origen, ordenados por cliente y fecha. Solo para comparación: NO usar para cargar.'],['Clientes sin correspondencia','En Todos J-R conservan el cliente original y se resaltan en amarillo con un comentario. Revisar Pendientes.'],['Fechas','Día 28 de cada mes; fecha convencional, no fecha real de emisión.'],['Pegado','Copiar A2:I hasta la última fila de Carga J-R y pegar en J:R de la base destino.'],['Criterio duplicados','Cliente normalizado + año + mes. Se excluye si ya existe, aunque cambie el consumo.'],['Datos','Se conservan ceros y textos. Solo se omiten vacíos.'],['Campos vacíos','Documento, límite, sobreconsumo e importes: sin datos en Hoja1; no se inventan.'],['Sobreconsumo','Columna O vacía. Usar pegado que omita blancos si hay fórmulas que deban preservarse.'],['Pendientes','No están en Carga J-R. Requieren revisión antes de una eventual carga.']]+[[k,v] for k,v in counts.items()]
 sheet('Instrucciones',['Concepto','Detalle'],notes)
 dest=Path(destination); dest.parent.mkdir(parents=True,exist_ok=True); out.save(dest)
 # Read-back validation of the actual output and reconciliation of all source cells.
 v=openpyxl.load_workbook(dest,read_only=True,data_only=False)
 actual=list(v['Carga J-R'].iter_rows(min_row=2,values_only=True)); seen=set()
 for row in actual:
  key=(norm(row[0]),row[1].year,row[1].month)
  if key in existing or key in seen: raise ValueError('Duplicado detectado al verificar la salida')
  seen.add(key)
  if row[1].day!=28 or row[1].year!=year: raise ValueError('Fecha de salida incorrecta')
  if any(row[i] is not None for i in (2,4,5,6,7,8)): raise ValueError('Campos reservados no vacíos')
 if counts['celdas_con_datos']!=len(rows)+len(excluded)+len(review)+counts['duplicados_internos_omitidos']: raise ValueError('No cuadran los registros de entrada y salida')
 if len(all_records)!=counts['celdas_con_datos'] or v['Todos J-R'].max_row-1!=len(all_records): raise ValueError('Faltan registros en Todos J-R')
 print(json.dumps(dict(counts),ensure_ascii=False,indent=2))
 print('PENDIENTES POR SUCURSAL')
 for x in sorted({(r[0],str(r[1]),str(r[3]),r[7]) for r in review}): print(x)
 print('VALIDACIÓN OK:',dest)
 v.close()
 w.close()
 return dict(counts)

def main():
 import argparse, sys
 base = Path(__file__).resolve().parent
 parser = argparse.ArgumentParser(description='Genera consumos faltantes para copiar a J:R sin modificar el Excel original.')
 parser.add_argument('--year', type=int, default=2026, help='Año de Hoja1 (por defecto: 2026)')
 parser.add_argument('--input', type=Path, default=base/'input', help='Carpeta de Excel de entrada')
 parser.add_argument('--output', type=Path, default=base/'output', help='Carpeta de resultados')
 args = parser.parse_args()
 if not 1900 <= args.year <= 9999:
  parser.error('El año debe estar entre 1900 y 9999')
 if args.input.resolve() == args.output.resolve():
  parser.error('Input y output deben ser carpetas distintas')
 if not args.input.is_dir():
  parser.error(f'No existe la carpeta de entrada: {args.input}')
 files = sorted(p for p in args.input.iterdir() if p.suffix.lower()=='.xlsx' and not p.name.startswith('~$'))
 if not files:
  print(f'No hay archivos .xlsx. Coloca tu Excel en: {args.input}', file=sys.stderr)
  return 1
 failures=0
 for source in files:
  destination=args.output/f'{source.stem}_faltantes_{args.year}.xlsx'
  print(f'Procesando: {source.name}', flush=True)
  if destination.exists():
   print(f'ERROR: ya existe {destination}. Muévelo o elimínalo antes de repetir.',file=sys.stderr)
   failures+=1
   continue
  try:
   process(source,destination,args.year)
  except Exception as error:
   print(f'ERROR en {source.name}: {error}',file=sys.stderr)
   failures+=1
 return 1 if failures else 0

if __name__=='__main__':
 raise SystemExit(main())
