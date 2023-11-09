
## 1. Importar e instalar librerias
# =================================
# =================================

### 1.1 Instalar
# ==============

# # Extraccion de datos
# !pip install pdfminer.six
# !pip install pdfquery
# !pip install PyPDF2
# !pip install "camelot-py[base]"
# !pip install ghostscript

# # Conversion a jpg
# !pip install pdf2image
# !sudo apt-get install python-poppler
# !sudo apt-get install poppler-utils

# # Conversion a txt
# !sudo add-apt-repository ppa:alex-p/tesseract-ocr-devel -y
# !sudo apt-get update
# !sudo apt-get install -y tesseract-ocr
# !sudo apt-get install tesseract-ocr-spa
# !pip install pytesseract
# !pip install pyyaml==5.1
# !pip install unidecode

### 1.2 Importar
# ==============

# Extraccion de datos
import os
import io
import camelot
import ghostscript
import PyPDF2
import re
from io import StringIO
import pandas as pd
import numpy as np
from pdfquery import PDFQuery
from pathlib import Path
import itertools
from itertools import chain
import matplotlib.pyplot as plt
from matplotlib import patches
import pdfminer
from pdfminer.high_level import extract_pages
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator 
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.layout import LTPage
from pdfminer.layout import LTTextBoxHorizontal
from pdfminer.layout import LTTextBoxVertical
from pdfminer.layout import LTTextLineHorizontal
from pdfminer.layout import LTTextLineVertical
from pdfminer.layout import LTTextContainer
from pdfminer.layout import LTChar
from pdfminer.layout import LTText
from pdfminer.layout import LTTextBox
from pdfminer.layout import LTAnno

# Conversion a jpg
from pdf2image import convert_from_path
from pathlib import Path
import PIL.Image
from glob import glob

# Conversion a txt
import json
import cv2
import random
import pytesseract
from pytesseract import Output
import matplotlib.pyplot as plt
from pdf2image import pdfinfo_from_path


## 2. Extracción de datos
# =======================
# =======================

def extract_page_layouts( file ):
    
    laparams        = LAParams()
    
    with open( file, mode = 'rb' ) as pdf_file:
        # print( 'Open document %s' % pdf_file.name )
        document    = PDFQuery( pdf_file ).doc
        
        if not document.is_extractable:
            raise PDFTextExtractionNotAllowed
            
        rsrcmgr     = PDFResourceManager()
        device      = PDFPageAggregator( rsrcmgr, laparams = laparams )
        interpreter = PDFPageInterpreter( rsrcmgr, device )
        layouts     = []
        
        for page in PDFPage.create_pages( document ):
            interpreter.process_page( page )
            layouts.append( device.get_result() )
            
    return layouts

# ============================================================================

def extract_characters( file ):

    document         = open( file, 'rb' )
    pdf              = PyPDF2.PdfFileReader( document )
    totalpages       = pdf.numPages
    characters       = []
    
    page_layouts     = extract_page_layouts( file )
    
    for page in range( totalpages ):
        current_page = page_layouts[ page ]

        for i, element in enumerate( current_page ):
            if isinstance( element, pdfminer.layout.LTTextBox ):
                for line in element:
                    if isinstance( line, pdfminer.layout.LTTextLineHorizontal ):
                        for char in line:
                            if isinstance( char, pdfminer.layout.LTChar ):
                                characters.append( (page, char ) )
                                
    return characters

# ==============================================================================

def order_characters( characters ):
    
    rows           = sorted( list (set( 800 *( p + 1 ) - c.bbox[ 1 ] for p, c in characters ) ), reverse = True )
    sorted_rows    = []
    
    for row in rows:
        sorted_row = sorted( [ c for p, c in characters if ( 800 *( p + 1 ) - c.bbox[ 1 ] ) == row ], key = lambda c: c.bbox[ 0 ] )
        sorted_rows.append( sorted_row )
        
    return sorted_rows

# ===============================================================================

def extract_info_from_table( archive ):
    
    filtered_rows = []
    table_rows    = []
    info          = []
    
    for row in archive:
        
        text_line        = "".join( [ c.get_text() for c in row ] )
        filtered_rows.append( row )
        
        if text_line.strip().startswith( 'DNI' ):
            
            colnames     = [ 'Civil', 'Penal', 'Admin.', 'Adm.ENT', 'Adm.PAS', 'Adm. ENT', 'Adm. PAS' ]
            start_column = {}
            
            for col in colnames:
                for i in range( len ( row ) ):
                    if text_line[ i: ].startswith( col ):
                        start_column[ col ] = row[ i ].bbox[ 0 ], row[ i ].bbox[ 1 ]
        
    for row in filtered_rows:
        line  = ''.join( [ c.get_text() for c in row ] )
        if line.startswith( 'DNI' ):
            continue
        match = re.match( r"\d{8}", line )
        if not match:
            continue
        table_rows.append( row )   
    
    for row in table_rows:
        text_line         = ''.join( [ c.get_text() for c in row ] )
        
        column_marks      = {}
        start_column_list = sorted( list ( start_column.items() ), key = lambda x: x[ 1 ][ 0 ] )
        
        for i, ( col, ( x, y ) ) in enumerate( start_column_list ):
            if i    == ( len( start_column_list ) - 1 ):
                x_f = 1e5
            else:
                x_f = start_column_list[ i + 1 ][ 1 ][ 0 ]
            
            marks   = []
            
            for element in row:
                if ( element.bbox[ 0 ] >= x ) and ( element.bbox[ 0 ] < x_f ):
                    marks.append( element.get_text() )
                    
            column_marks[ col ] = marks
        
        info.append( ( text_line, column_marks ) )
    info = info[ :: - 1 ]
        
    df   = pd.DataFrame( info, columns = [ 'nombre', 'columnas' ] )
    return df

# =================================================================================

def fix_tables( df ):
    
    df[ 'dni' ]      = df[ 'nombre' ].str[ :8 ]
    df[ 'nombres' ]  = df[ 'nombre' ].str[ 8: ]
    df[ 'personas' ] = df[ 'nombres' ].map( lambda x: x.lstrip( '+-' ).rstrip( 'X' ) )  
    #df[ 'dni' ]      = '_' + df[' dni' ].astype( str )
    df               = df[ [ 'dni','columnas', 'personas' ] ]
    
    return df

# =================================================================================

def handle_missing_info( df ):
    
    if df.empty:
        new_row = { 'dni': '.', 
                    'nombres': '.',
                    'columnas': '.',
                    'personas': '.',
                    'doc_name': '.' }
        
        df = pd.concat( [ df, pd.DataFrame( [ new_row ] ) ], ignore_index = True )
        return df
    
    else:
        return df.copy()
    
# =================================================================================

def create_responsibility_vars( df ):

    df[ 'test' ]    = df[ 'columnas' ].astype(str)
    df[ 'test' ]    = df[ 'test' ].str.replace( r"\[]", "0", regex=  True )
    df[ 'test' ]    = df[ 'test' ].str.replace( r"\['X']", "1", regex = True )
    df[ 'test' ]    = df[ 'test' ].str.replace( r"Adm. ENT|Adm.ENT", "Adm_ENT", regex = True )
    df[ 'test' ]    = df[ 'test' ].str.replace( r"Adm. PAS|Adm.PAS", "Adm_PAS", regex = True )
    df[ 'test' ]    = df[ 'test' ].str.replace( r"Admin.", "Admin", regex = True )

    pattern_civil   = r"'Civil':\s*(\d+)"
    pattern_penal   = r"'Penal':\s*(\d+)"
    pattern_admin   = r"'Admin':\s*(\d+)"
    pattern_adm_ent = r"'Adm_ENT':\s*(\d+)"
    pattern_adm_pas = r"'Adm_PAS':\s*(\d+)"

    df['civil']     = df[ 'test' ].str.extract( pattern_civil, expand = False )
    df['penal']     = df[ 'test' ].str.extract( pattern_penal, expand = False )
    df['admin']     = df[ 'test' ].str.extract( pattern_admin, expand = False )
    df['adm_ent']   = df[ 'test' ].str.extract( pattern_adm_ent, expand = False )
    df['adm_pas']   = df[ 'test' ].str.extract( pattern_adm_pas, expand = False )

    df              = df.drop( [ 'test' ], axis = 1 )
    
    return df

# ==================================================================================

def extract_table( archivo ):
    
    characters  = extract_characters( archivo )
    sorted_rows = order_characters( characters )
    table       = extract_info_from_table( sorted_rows )
    table       = fix_tables( table )
    table       = handle_missing_info( table )
    table       = create_responsibility_vars( table )
    
    return table    

# =================================================================================
# =================================================================================

def get_ubigeo(archivo, modalidad ):
    
    reader          = camelot.read_pdf( archivo, pages = '1', flavor = 'lattice', line_scale = 40 )
    tabla           = reader[ 0 ].df.T

    if modalidad   == 'AC':
        columns     = [ 'numero_informe', 'titulo_informe/asunto', 'objetivo', 'entidad_auditada', 
                        'monto_auditado', 'monto_examinado', 'ubigeo', 'fecha_emision_informe', 'unidad_emite_informe' ]
        
    elif modalidad == 'SEHPI':
        columns     = [ 'numero_informe', 'titulo_informe/asunto', 'objetivo', 'entidad_auditada', 
                       'monto_objeto_servicio', 'ubigeo', 'fecha_emision_informe', 'unidad_emite_informe' ]

    tabla.columns        = columns
    tabla                = tabla.iloc[ 1: ] 
    
    tabla[ 'distrito' ]  = tabla[ 'ubigeo' ].str.extract( 'Distrito:\s*(.+)', re.MULTILINE | re.DOTALL ).fillna( '' )
    tabla[ 'provincia' ] = tabla[ 'ubigeo' ].str.extract( 'Provincia:\s*(.+)', re.MULTILINE | re.DOTALL ).fillna( '' )
    tabla[ 'region' ]    = tabla[ 'ubigeo' ].str.extract( 'Región:\s*(.+)', re.MULTILINE | re.DOTALL ).fillna( '' )
    
    tabla                = tabla.drop( [ 'ubigeo' ], axis = 1 ) 
    tabla                = tabla.apply( lambda x: ' '.join( x.dropna().astype( str ) ) ).to_frame().T
    return tabla

# =================================================================================

# def get_text( archivo, modalidad ):

#     tabla              = camelot.read_pdf( archivo, pages = 'all', flavor = 'lattice' )
#     illegal_characters = re.compile( r'[\000-\010]|[\013-\014]|[\016-\037]' )
    
#     if modalidad      == 'AC':
#         first_column   = 'observaciones'
#     elif modalidad    == 'SEHPI':
#         first_column   = 'argumentos_de_hecho'
        
#     if tabla.n != 4:
#         df_appended     = pd.DataFrame()
        
#         for i in range( 2, tabla.n, 1 ):
#             df          = tabla[ i ].df
#             df_appended = pd.concat( [ df_appended, df ], ignore_index = True )
            
#         df_appended     = df_appended.replace( '', np.nan ).ffill().bfill()
#         df_appended     = df_appended.rename( columns = { 1: 'text' } )
#         df_grouped      = pd.DataFrame( df_appended.groupby( 0 )[ 'text' ].apply( lambda x: '\n'.join( x ) ) ).T
#         df_grouped      = df_grouped.rename( columns = { '1': first_column, '2': 'recomendaciones', '3': 'funcionarios' } )
#         df_grouped      = df_grouped.reset_index( drop = True )
#         df_grouped      = df_grouped.applymap( lambda x: illegal_characters.sub( r'', x ) if isinstance( x, str ) else x )
#         df_grouped      = df_grouped.drop( [ 'funcionarios' ], axis = 1 )
            
#     elif tabla.n == 4:
        
#         tabla2    = tabla[ 2 ].df
#         tabla3    = tabla[ 3 ].df
#         cols2     = set( tabla2.columns )
#         cols3     = set( tabla3.columns )
        
#         if cols2 == cols3:
#             df    = pd.concat( [ tabla2, tabla3 ], ignore_index = True )
#         else:
#             df    = tabla3
            
#         df            = df.replace( '', np.nan ).ffill().bfill()
#         df            = df.rename( columns = { 1: 'text' } )
#         df_grouped    = pd.DataFrame( df.groupby( 0 )[ 'text' ].apply( lambda x: '\n'.join( x ) ) ).T
#         df_grouped    = df_grouped.rename( columns = { '1': first_column, '2': 'recomendaciones', '3': 'funcionarios' } )
#         df_grouped    = df_grouped.reset_index( drop = True )
#         df_grouped    = df_grouped.applymap( lambda x: illegal_characters.sub( r'', x ) if isinstance( x, str ) else x )
#         df_grouped    = df_grouped.drop( [ 'funcionarios' ], axis = 1 )
        
#     return df_grouped


def get_text( archivo, modalidad ):

    tabla              = camelot.read_pdf( archivo, pages = 'all', flavor = 'lattice' )
    illegal_characters = re.compile( r'[\000-\010]|[\013-\014]|[\016-\037]' )
    
    if modalidad      == 'AC':
        first_column   = 'observaciones'
    elif modalidad    == 'SEHPI':
        first_column   = 'argumentos_de_hecho'
        
    if tabla.n != 4:
        df_appended     = pd.DataFrame()
        
        for i in range( 2, tabla.n, 1 ):
            df          = tabla[ i ].df
            df_appended = pd.concat( [ df_appended, df ], ignore_index = True )
            
        df_appended        = df_appended.replace( '', np.nan ).ffill().bfill()
        df_appended        = df_appended.rename( columns = { 1: 'text' } )
        num_rows           = len( df_appended )

        try: 

            idx_recomendations = df_appended[ df_appended[ 'text' ].str.contains( 'recomendaciones', case = False ) ].index[ 0 ]
            idx_last_row       = num_rows - 1 if idx_recomendations < num_rows - 1 else num_rows   

            first_column_text  = df_appended.loc[ 0:idx_recomendations-1, 'text' ].str.cat( sep = '\n' )
            recomendations     = df_appended.loc[ idx_recomendations:idx_last_row-1, 'text' ].str.cat( sep = '\n' )

            df_appended        = pd.DataFrame( {first_column: [ first_column_text ], 'recomendaciones': [ recomendations ] } )

        except IndexError:

            first_column_text = df_appended.loc[ 0:num_rows, 'text' ].str.cat( sep = '\n' )
            recomendations    = f'NOTA QLAB. NO SE ENCONTRARON RECOMENDACIONES. REVÍSESE LA COLUMNA "{ first_column }"'
            df_appended       = pd.DataFrame( {first_column: [ first_column_text ], 'recomendaciones': [ recomendations ] } )
        
                   
    elif tabla.n == 4:
        
        tabla2    = tabla[ 2 ].df
        tabla3    = tabla[ 3 ].df
        cols2     = set( tabla2.columns )
        cols3     = set( tabla3.columns )
        
        if cols2 == cols3:
            df_appended    = pd.concat( [ tabla2, tabla3 ], ignore_index = True )
        else:
            df_appended    = tabla3
            
        df_appended        = df_appended.replace( '', np.nan ).ffill().bfill()
        df_appended        = df_appended.rename( columns = { 1: 'text' } )
        num_rows           = len( df_appended )

        try: 

            idx_recomendations = df_appended[ df_appended[ 'text' ].str.contains( 'recomendaciones', case = False ) ].index[ 0 ]
            idx_last_row       = num_rows - 1 if idx_recomendations < num_rows - 1 else num_rows   

            first_column_text  = df_appended.loc[ 0:idx_recomendations-1, 'text' ].str.cat( sep = '\n' )
            recomendations     = df_appended.loc[ idx_recomendations:idx_last_row-1, 'text' ].str.cat( sep = '\n' )

            df_appended        = pd.DataFrame( {first_column: [ first_column_text ], 'recomendaciones': [ recomendations ] } )

        except IndexError:

            first_column_text = df_appended.loc[ 0:num_rows, 'text' ].str.cat( sep = '\n' )
            recomendations    = f'NOTA QLAB. NO SE ENCONTRARON RECOMENDACIONES. REVÍSESE LA COLUMNA "{ first_column }"'
            df_appended       = pd.DataFrame( {first_column: [ first_column_text ], 'recomendaciones': [ recomendations ] } )
        
    return df_appended

# =======================================================================================

def extract_info( archivo, modalidad ):
    
    ubigeo               = get_ubigeo( archivo, modalidad )
    texto                = get_text( archivo, modalidad )
    
    tabla                = pd.concat( [ ubigeo, texto ], axis = 1 )
    tabla[ 'modalidad' ] = modalidad
    
    return tabla

# =======================================================================================

def extract_all( archivo, modalidad ):

    tabla_resp          = extract_table( archivo )
    tabla_info          = extract_info( archivo, modalidad )
    tabla_info_dup      = pd.concat( [ tabla_info ] * len( tabla_resp ), ignore_index = True )
    
    data                = pd.concat( [ tabla_info_dup, tabla_resp ], axis = 1 )
    data[ 'modalidad' ] = modalidad

    return data    


## 3. Conversiones
# ================
# ================

### 3.1. Convertir a jpg
# ======================

# def convert_1p_to_jpg( input_path, output_path, modalidad ):
    
#     PIL.Image.MAX_IMAGE_PIXELS = 933120000
    
#     with open( f'{ modalidad }_to_jpg.txt', 'w+' ) as file:
        
#         for i, input_element in enumerate( input_path ):
            
#             input_element_name = input_element[ input_element.rindex( '/' ) + 1: ]
#             input_element_name = os.path.splitext( input_element_name )[ 0 ]
            
#             converted_input_elements = convert_from_path( input_element, 500, first_page = 1, last_page = 1 )
#             os.makedirs( output_path, exist_ok = True )
            
#             for converted_input_element in converted_input_elements:
                
#                 output_element_path = os.path.join( output_path, f'{ input_element_name }.jpg' )
#                 converted_input_element.save( output_element_path, 'JPEG' )
                
#                 line = f'{ output_element_path }\t{ input_element }\n'
#                 file.write( line )



def convert_1p_to_jpg( input_path, output_path, txt_output_path, modalidad ):
    PIL.Image.MAX_IMAGE_PIXELS = 933120000
    
    txt_file_path = os.path.join( txt_output_path, f'{ modalidad }_to_jpg.txt' )
    
    with open( txt_file_path, 'w+' ) as file:
        for i, input_element in enumerate( input_path ):
            try:
                input_element_name       = input_element[ input_element.rindex( '/' ) + 1: ]
                input_element_name       = os.path.splitext( input_element_name )[ 0 ]
    
                converted_input_elements = convert_from_path( input_element, 500, first_page = 1, last_page = 1 )
                os.makedirs( output_path, exist_ok = True )
                
                for converted_input_element in converted_input_elements:
                    output_element_path = os.path.join( output_path, f'{ input_element_name }.jpg' )
                    converted_input_element.save( output_element_path, 'JPEG' )

                    line = f'Success converting: { output_element_path }\t{ input_element }\n'
                    file.write( line )
                    
            except Exception as e:
                error_line = f'Error converting: { input_element }\n'
                file.write( error_line )

                
### 3.2. Convertir a txt
# ======================

def convert_1p_to_text( input_path, output_path, txt_output_path, modalidad ):
    
    txt_file_path = os.path.join( txt_output_path, f'{ modalidad }_to_text.txt' )
    
    with open( txt_file_path, 'w+' ) as file:
        
        for i, input_element in enumerate( input_path ):
            
            try:            
                input_element_name      = input_element[ input_element.rindex( '/' ) + 1: ]
                input_element_name      = os.path.splitext( input_element_name )[ 0 ]

                image                   = cv2.imread( input_element )
                rgb                     = cv2.cvtColor( image, cv2.COLOR_BGR2RGB )
                hImg, WImg, _           = image.shape
                results                 = pytesseract.image_to_string( image, lang = 'spa' )

                converted_input_element = [ line for line in results.split( '\n' ) if line.strip() != '' ]
                output_element_path     = os.path.join( output_path, f'{ input_element_name }.txt' )

                with open( output_element_path, 'w+' ) as text:
                    text.write( '\n'.join( converted_input_element ) )

                line = f'Success converting: { output_element_path }\t{ input_element }\n'
                file.write( line )  
            
            except Exception as e:
                error_line = f'Error converting: { input_element }\n'
                file.write( error_line )          


### 3.3. Extraer fechas de los informes
# =====================================

def extract_dates( input_path, output_path, txt_output_path, modalidad, regex ):
    
    txt_to_dates_file_path = os.path.join( txt_output_path, f'{ modalidad }_to_dates.txt' )
    
    with open( txt_to_dates_file_path, 'w+' ) as file:

        not_recognized_date = []

        for i, input_element in enumerate( input_path ):

            input_element_name      = input_element[ input_element.rindex( '/' ) + 1: ]
            input_element_name      = os.path.splitext( input_element_name )[ 0 ]
            
            with open( input_element ) as date:
                
                reader     = date.read()
                
            reader_matches = list ( re.finditer( regex, reader, re.MULTILINE ) )
            
            if len( reader_matches ) == 0:
                not_recognized_date.append( input_element )
                
            else:
                reader_matches             = reader_matches[ -1 ]
                extracted_dates            = reader_matches.groupdict()
                extracted_dates            = pd.DataFrame( extracted_dates, index = [ 0 ] )
                extracted_dates[ 'txt_name' ] = input_element_name
                output_element_path        = os.path.join( output_path, f'{ input_element_name }.xlsx' )
                extracted_dates.to_excel( output_element_path )

            line = f'{ output_element_path }\t{ input_element }\n'
            file.write( line )       

        txt_not_recognized_file_path = os.path.join( txt_output_path, f'{ modalidad }_not_recognized_dates.txt' )
        
        with open( txt_not_recognized_file_path, 'w+' ) as nrd:
            for element in not_recognized_date:
                nrd.write( f'{ element }\n' ) 

# =========================================================================================

def append_dates( input_path, output_path, modalidad ):
    
    input_elements = glob( input_path )
    tables         = []
    
    for input_element in input_elements:
        
        table       = pd.read_excel( input_element )
        table       = table.fillna( '.' )

        table.loc[ table.month1 == '.', 'month1' ] = table.month2
        table.loc[ table.year1 == '.', 'year1' ] = table.year2

        tables.append( table )
        
    tables_appended                     = pd.concat( tables, axis = 0 )
    
    tables_appended[ 'numero_informe' ] = tables_appended[ 'txt_name' ].str.replace( '-informe', '' )
    tables_appended                     = tables_appended.reset_index( drop = True )

    tables_appended[ 'day1' ]           = tables_appended[ 'day1' ].astype( str )
    tables_appended[ 'month1' ]         = tables_appended[ 'month1' ].str.lower().astype( str )
    tables_appended[ 'year1' ]          = tables_appended[ 'year1' ].astype( str )
  
    tables_appended[ 'day2' ]           = tables_appended[ 'day2' ].astype( str )
    tables_appended[ 'month2' ]         = tables_appended[ 'month2' ].str.lower().astype( str )
    tables_appended[ 'year2' ]          = tables_appended[ 'year2' ].astype( str )

    meses = {
        'enero': 'January',
        'febrero': 'February',
        'marzo': 'March',
        'abril': 'April',
        'mayo': 'May',
        'junio': 'June',
        'julio': 'July',
        'agosto': 'August',
        'septiembre': 'September',
        'setiembre': 'September',
        'octubre': 'October',
        'noviembre': 'November',
        'diciembre': 'December'
    }


    tables_appended[ 'inicio' ] = tables_appended[ 'day1' ].astype(str) + '/' + tables_appended[ 'month1' ].map( meses ) + '/' + tables_appended[ 'year1' ].astype( str )
    tables_appended[ 'final' ]  = tables_appended[ 'day2' ].astype(str) + '/' + tables_appended[ 'month2' ].map( meses ) + '/' + tables_appended[ 'year2' ].astype( str )

    tables_appended[ 'inicio' ] = pd.to_datetime( tables_appended[ 'inicio' ], format = '%d/%B/%Y' ).dt.strftime( '%d/%m/%Y' )
    tables_appended[ 'final' ]  = pd.to_datetime( tables_appended[ 'final' ], format = '%d/%B/%Y' ).dt.strftime( '%d/%m/%Y' )

    tables_appended             = tables_appended.drop( columns = [ 'day1', 'month1', 'year1', 'day2', 'month2', 'year2', 'Unnamed: 0', 'txt_name' ] )

    output_element_path          = os.path.join( output_path, f'{ modalidad }_fechas.xlsx' )
    tables_appended.to_excel( output_element_path, index = False )
        
    return tables_appended