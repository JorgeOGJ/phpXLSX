<?php

/**
 * Clase para generar un archivo xls de manera óptima basandose en el estándar de
 * openxml para excel
 */
class build_excel
{

    private $_nombreArchivo = '';
    private $_zip;/*Se trabaja con un archivo zip dado que los archivos xlsx son un zip con una estructura*/
    /*
     * Las siguientes variables contienen las cadenas que al final son convertidas en archivos
     * xml y que son parte de la estructura interna del xlsx
     */
    private $_app = '';
    private $_core = '';
    private $_xlrels = '';
    private $_contentType = '';
    private $_rels = '';
    private $_workbook = '';
    private $_styles = '';
    private $_theme = '';
    private $_sheet = '';
    private $_sharedStrings = '';
    
    /*
     * Esta variable contiene la información de todas las filas del documento
     */
    private $_dataSheet = '';
    /*
     * El formato openxml require de la creacion de una LookUp Table donde esten
     * contenidas las cadenas para cada celda del xlsx
     */
    private $_LUT = array();
    private $_k = 0;
    private $_di = 'A1';
    private $_df = '';
    private $_contador = 0;
    private $_LUTindex = 0;
    private $_encabezados = array();
    private $_headers = '';
    private $_titulo = '';

    /**
     * Nombre de archivo a ser generado
     * @param string ruta para generar archivo
     * @return archivo a ser generado
     */
    function build_excel($file)
    {

        return $this->open($file);
    }

    /**
     * Función interna para crear el recurso de archivo
     * @param string $file ruta de archivo
     * @return recurso de archivo creado
     */
    private function open($file)
    {
        $this->_zip = new Zip();
        $this->_zip->setZipFile($file);
        return $this->_zip;
    }

    /**
     * Función que finaliza el archivo zip
     * archivo xls
     * @return type
     */
    private function close()
    {
        $this->_zip->finalize();
        return;
    }

    /**
     * Función que escribe en el archivo los nombres de las cabeceras dadas para 
     * las columnas del documento
     * @param array $array Arreglo que contiene las cabeceras para las columnas 
     */
    private function crear_encabezados()
    {
        for ($i = 0; $i < count($this->_encabezados); $i++)
        {
            $this->_headers.='<col min="' . ($i + 1) . '" max="' . ($i + 1) . '" width="36.5703125" bestFit="1" customWidth="1"/>';
            $this->_LUT[$this->_k] = $this->_encabezados[$i];
            $this->_k++;
        }
    }

    /***
     * Función que crea el titulo en caso de existir para la hoja de datos
     */
    private function crear_titulo()
    {
        $this->_LUT[$this->_k] = $this->_titulo;
        $this->_k++;
        $this->_dataSheet.='<row r="1" spans="1:13" ht="25.5" x14ac:dyDescent="0.2">
                              <c r="A1" s="1" t="s">
                                <v>0</v>
                              </c>
                            </row>';
        $this->_contador++;
    }

    /**
     * Función que escribe en el documento una fila completa dado un array de datos
     * @param array $array
     */
    private function escribir_fila($array)
    {
        if (!empty($array) && !is_null($array))
        {
            $this->_dataSheet.='<row r="' . $this->_contador . '" spans="1:' . count($this->_encabezados) . '" x14ac:dyDescent="0.2">';
            $cell = 'A';
            foreach ($array as $dato)
            {
                $this->_LUT[$this->_k] = $dato;
                $this->_dataSheet .= '<c r="' . $cell . $this->_contador . '" s="1" t="s">
                     <v>' . $this->_LUTindex . '</v>
                    </c>';
                $this->_LUTindex++;
                $this->_k++;
                $cell = chr(ord($cell) + 1);
            }
            $this->_dataSheet.='</row>';
        }
    }

    /**
     * Función que obtiene el archivo creado, la ruta o descarga directamente al
     * navegador el archivo dependiendo de los parametros de entrada
     * @param boolean $web Bandera para indicar que se descargue en el navegador
     * @param boolean $ruta Bandera para indicar que se regrese la ruta del archivo
     * @return mixed Regresa un recurso, ruta o el archivo directamente al navegador
     */
    public function getFile($web = true, $ruta = false)
    {
        $this->close();
        if (file_exists($this->_nombreArchivo))
        {
            if ($web)
            {

                header('Content-Description: File Transfer');
                header('Content-Type: application/octet-stream');
                header('Content-Disposition: attachment; filename="' . basename($this->_nombreArchivo) . '"');
                header('Content-Transfer-Encoding: binary');
                header('Connection: Keep-Alive');
                header('Expires: 0');
                header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
                header('Pragma: public');
                header('Content-Length: ' . filesize($this->_nombreArchivo));
                readfile($this->_nombreArchivo);
            }
            if ($ruta)
            {
                return $this->_nombreArchivo;
            }
            else
            {
                return $this->_zip;
            }
        }
    }

    public function createFromResultSet($result, $titulo, $encabezado)
    {
        $this->_contador = 1;
        $this->_LUTindex = 1;

        if (!empty($titulo))
        {
            $this->_titulo = $titulo;
            $this->crear_titulo();
        }
        else
        {
            $this->_titulo = 'Consulta';
            $this->crear_titulo();
        }

        if (is_null($encabezado))
        {
            for ($i = 0; $i < $result->columnCount(); $i++)
            {
                $nombre = $result->getColumnMeta($i);

                array_push($this->_encabezados, $nombre['name']);
            }
            $this->crear_encabezados();
        }
        else
        {
            $this->_encabezados = $encabezado;
            $this->crear_encabezados();
        }

        $hora_inicio = date('H:i:s');
        error_log('INICIO ' . date('Y-m-d H:i:s') . ' ' . (memory_get_peak_usage(true) / 1024 / 1024) . ' MB');

        while ($row = $result->fetch(PDO::FETCH_NUM, PDO::FETCH_ORI_NEXT, PDO::FETCH_ORI_REL))
        {
            $this->escribir_fila($row);
            $this->_contador++;
        }
        $this->crear_sharedStrings();
        $this->crear_sheet();
        $this->crear_estructura();

        error_log('REGISTROS ' . $this->_contador);
        $hora_final = date('H:i:s');
        error_log('TIEMPO ESTIMADO DE PROCESO: ' . date('i:s', strtotime($hora_final) - strtotime($hora_inicio)));
        error_log('FINAL ' . date('Y-m-d H:i:s') . ' ' . (memory_get_peak_usage(true) / 1024 / 1024) . ' MB');
    }

    private function crear_sharedStrings()
    {
        $sSD = '';
        for ($l = 0; $l < count($this->_LUT); $l++)
        {
            $sSD.='<si><t>' . $this->_LUT[$l] . '</t></si>' . "\n";
        }

        $this->_sharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                                    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                                    ' . $sSD . '
                                    </sst>';
    }

    private function crear_sheet()
    {
        $this->_df = chr(ord('A') + (count($this->_encabezados) - 1));
        $this->_df.=$this->_contador;

        $dimension = $this->_di . ':' . $this->_df;

        $this->_sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <dimension ref="' . $dimension . '"/>
    <sheetViews>
    <sheetView tabSelected="1" topLeftCell="A2" workbookViewId="0"/>
    </sheetViews>
    <sheetFormatPr baseColWidth="10" defaultColWidth="9.140625" defaultRowHeight="12.75" x14ac:dyDescent="0.2"/>
    <cols>
    ' .
                $this->_headers
                . '
    </cols>

    <sheetData>
    ' .
                $this->_dataSheet
                . '
    </sheetData>

    <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
    </worksheet>';
    }

    private function crear_estructura()
    {
        $this->_app = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime></Properties>';
        $this->_core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">' . gmDate("Y-m-d\TH:i:s\Z") . '</dcterms:created><cp:revision>0</cp:revision></cp:coreProperties>';
        $this->_xlrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>';
        $this->_contentType = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
        $this->_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

        $this->_workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9302"/><workbookPr defaultThemeVersion="124226"/>
<bookViews><workbookView xWindow="120" yWindow="135" windowWidth="10005" windowHeight="10005"/></bookViews>
<sheets><sheet name="Reporte" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="0"/></workbook>';
        $this->_styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" 
xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
<fonts count="20" x14ac:knownFonts="1"><font><sz val="10"/><name val="Arial"/></font><font><sz val="11"/><color theme="1"/><name val="Calibri"/>
<family val="2"/><scheme val="minor"/></font><font><b/><sz val="18"/><color theme="3"/><name val="Cambria"/><family val="2"/><scheme val="major"/>
</font><font><b/><sz val="15"/><color theme="3"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="13"/>
<color theme="3"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="3"/><name val="Calibri"/>
<family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FF9C6500"/>
<name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FF3F3F76"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color rgb="FF3F3F3F"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color rgb="FFFA7D00"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FFFA7D00"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="0"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><i/><sz val="11"/><color rgb="FF7F7F7F"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color theme="0"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font></fonts><fills count="34"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC99"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFA5A5A5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFCC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="5" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="5" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="5" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="6"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="6" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="6" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="6" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="7"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="7" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="7" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="7" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="8"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="8" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="8" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="8" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="9"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="9" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="9" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="9" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF205081"/><bgColor rgb="FF003366"/></patternFill></fill></fills><borders count="10"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thick"><color theme="4"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="thick"><color theme="4" tint="0.499984740745262"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="medium"><color theme="4" tint="0.39997558519241921"/></bottom><diagonal/></border><border><left style="thin"><color rgb="FF7F7F7F"/></left><right style="thin"><color rgb="FF7F7F7F"/></right><top style="thin"><color rgb="FF7F7F7F"/></top><bottom style="thin"><color rgb="FF7F7F7F"/></bottom><diagonal/></border><border><left style="thin"><color rgb="FF3F3F3F"/></left><right style="thin"><color rgb="FF3F3F3F"/></right><top style="thin"><color rgb="FF3F3F3F"/></top><bottom style="thin"><color rgb="FF3F3F3F"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="double"><color rgb="FFFF8001"/></bottom><diagonal/></border><border><left style="double"><color rgb="FF3F3F3F"/></left><right style="double"><color rgb="FF3F3F3F"/></right><top style="double"><color rgb="FF3F3F3F"/></top><bottom style="double"><color rgb="FF3F3F3F"/></bottom><diagonal/></border><border><left style="thin"><color rgb="FFB2B2B2"/></left><right style="thin"><color rgb="FFB2B2B2"/></right><top style="thin"><color rgb="FFB2B2B2"/></top><bottom style="thin"><color rgb="FFB2B2B2"/></bottom><diagonal/></border><border><left/><right/><top style="thin"><color theme="4"/></top><bottom style="double"><color theme="4"/></bottom><diagonal/></border></borders><cellStyleXfs count="42"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="3" fillId="0" borderId="1" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="4" fillId="0" borderId="2" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="5" fillId="0" borderId="3" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="5" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="6" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="7" fillId="3" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="8" fillId="4" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="9" fillId="5" borderId="4" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="10" fillId="6" borderId="5" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="11" fillId="6" borderId="4" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="12" fillId="0" borderId="6" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="13" fillId="7" borderId="7" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="14" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="8" borderId="8" applyNumberFormat="0" applyFont="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="15" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="16" fillId="0" borderId="9" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="9" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="10" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="11" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="12" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="13" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="14" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="15" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="16" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="17" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="18" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="19" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="20" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="21" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="22" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="23" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="24" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="25" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="26" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="27" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="28" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="29" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="30" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="31" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="32" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="18" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1"/></xf><xf numFmtId="0" fontId="19" fillId="33" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs><cellStyles count="42"><cellStyle name="20% - Énfasis1" xfId="19" builtinId="30" customBuiltin="1"/><cellStyle name="20% - Énfasis2" xfId="23" builtinId="34" customBuiltin="1"/><cellStyle name="20% - Énfasis3" xfId="27" builtinId="38" customBuiltin="1"/><cellStyle name="20% - Énfasis4" xfId="31" builtinId="42" customBuiltin="1"/><cellStyle name="20% - Énfasis5" xfId="35" builtinId="46" customBuiltin="1"/><cellStyle name="20% - Énfasis6" xfId="39" builtinId="50" customBuiltin="1"/><cellStyle name="40% - Énfasis1" xfId="20" builtinId="31" customBuiltin="1"/><cellStyle name="40% - Énfasis2" xfId="24" builtinId="35" customBuiltin="1"/><cellStyle name="40% - Énfasis3" xfId="28" builtinId="39" customBuiltin="1"/><cellStyle name="40% - Énfasis4" xfId="32" builtinId="43" customBuiltin="1"/><cellStyle name="40% - Énfasis5" xfId="36" builtinId="47" customBuiltin="1"/><cellStyle name="40% - Énfasis6" xfId="40" builtinId="51" customBuiltin="1"/><cellStyle name="60% - Énfasis1" xfId="21" builtinId="32" customBuiltin="1"/><cellStyle name="60% - Énfasis2" xfId="25" builtinId="36" customBuiltin="1"/><cellStyle name="60% - Énfasis3" xfId="29" builtinId="40" customBuiltin="1"/><cellStyle name="60% - Énfasis4" xfId="33" builtinId="44" customBuiltin="1"/><cellStyle name="60% - Énfasis5" xfId="37" builtinId="48" customBuiltin="1"/><cellStyle name="60% - Énfasis6" xfId="41" builtinId="52" customBuiltin="1"/><cellStyle name="Buena" xfId="6" builtinId="26" customBuiltin="1"/><cellStyle name="Cálculo" xfId="11" builtinId="22" customBuiltin="1"/><cellStyle name="Celda de comprobación" xfId="13" builtinId="23" customBuiltin="1"/><cellStyle name="Celda vinculada" xfId="12" builtinId="24" customBuiltin="1"/><cellStyle name="Encabezado 4" xfId="5" builtinId="19" customBuiltin="1"/><cellStyle name="Énfasis1" xfId="18" builtinId="29" customBuiltin="1"/><cellStyle name="Énfasis2" xfId="22" builtinId="33" customBuiltin="1"/><cellStyle name="Énfasis3" xfId="26" builtinId="37" customBuiltin="1"/><cellStyle name="Énfasis4" xfId="30" builtinId="41" customBuiltin="1"/><cellStyle name="Énfasis5" xfId="34" builtinId="45" customBuiltin="1"/><cellStyle name="Énfasis6" xfId="38" builtinId="49" customBuiltin="1"/><cellStyle name="Entrada" xfId="9" builtinId="20" customBuiltin="1"/><cellStyle name="Incorrecto" xfId="7" builtinId="27" customBuiltin="1"/><cellStyle name="Neutral" xfId="8" builtinId="28" customBuiltin="1"/><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Notas" xfId="15" builtinId="10" customBuiltin="1"/><cellStyle name="Salida" xfId="10" builtinId="21" customBuiltin="1"/><cellStyle name="Texto de advertencia" xfId="14" builtinId="11" customBuiltin="1"/><cellStyle name="Texto explicativo" xfId="16" builtinId="53" customBuiltin="1"/><cellStyle name="Título" xfId="1" builtinId="15" customBuiltin="1"/><cellStyle name="Título 1" xfId="2" builtinId="16" customBuiltin="1"/><cellStyle name="Título 2" xfId="3" builtinId="17" customBuiltin="1"/><cellStyle name="Título 3" xfId="4" builtinId="18" customBuiltin="1"/><cellStyle name="Total" xfId="17" builtinId="25" customBuiltin="1"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>';

        $this->_theme = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Tema de Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';


        $this->_zip->addFile($this->_contentType, '[Content_Types].xml');

        $this->_zip->addDirectory('docProps');
        $this->_zip->addFile($this->_app, 'docProps/app.xml');
        $this->_zip->addFile($this->_core, 'docProps/core.xml');

        $this->_zip->addDirectory('xl');
        $this->_zip->addDirectory('xl/worksheets');
        $this->_zip->addFile($this->_sheet, 'xl/worksheets/sheet1.xml');

        $this->_zip->addFile($this->_sharedStrings, 'xl/sharedStrings.xml');
        $this->_zip->addFile($this->_workbook, 'xl/workbook.xml');
        $this->_zip->addFile($this->_styles, 'xl/styles.xml');

        $this->_zip->addDirectory('xl/theme');
        $this->_zip->addFile($this->_theme, 'xl/theme/theme1.xml');

        $this->_zip->addDirectory('xl/_rels');
        $this->_zip->addFile($this->_xlrels, 'xl/_rels/workbook.xml.rels');

        $this->_zip->addDirectory('_rels');
        $this->_zip->addFile($this->_rels, '_rels/.rels');
    }

}

/**
 * Class to create and manage a Zip file.
 *
 * Initially inspired by CreateZipFile by Rochak Chauhan  www.rochakchauhan.com (http://www.phpclasses.org/browse/package/2322.html)
 * and
 * http://www.pkware.com/documents/casestudies/APPNOTE.TXT Zip file specification.
 *
 * License: GNU LGPL, Attribution required for commercial implementations, requested for everything else.
 *
 * @author A. Grandt <php@grandt.com>
 * @copyright 2009-2014 A. Grandt
 * @license GNU LGPL 2.1
 * @link http://www.phpclasses.org/package/6110
 * @link https://github.com/Grandt/PHPZip
 * @version 1.61
 */
class Zip
{

    const VERSION = 1.61;
    const ZIP_LOCAL_FILE_HEADER = "\x50\x4b\x03\x04"; // Local file header signature
    const ZIP_CENTRAL_FILE_HEADER = "\x50\x4b\x01\x02"; // Central file header signature
    const ZIP_END_OF_CENTRAL_DIRECTORY = "\x50\x4b\x05\x06\x00\x00\x00\x00"; //end of Central directory record
    const EXT_FILE_ATTR_DIR = 010173200020;  // Permission 755 drwxr-xr-x = (((S_IFDIR | 0755) << 16) | S_DOS_D);
    const EXT_FILE_ATTR_FILE = 020151000040; // Permission 644 -rw-r--r-- = (((S_IFREG | 0644) << 16) | S_DOS_A);
    const ATTR_VERSION_TO_EXTRACT = "\x14\x00"; // Version needed to extract
    const ATTR_MADE_BY_VERSION = "\x1E\x03"; // Made By Version
    // UID 1000, GID 0
    const EXTRA_FIELD_NEW_UNIX_GUID = "\x75\x78\x0B\x00\x01\x04\xE8\x03\x00\x00\x04\x00\x00\x00\x00";

    // Unix file types
    const S_IFIFO = 0010000; // named pipe (fifo)
    const S_IFCHR = 0020000; // character special
    const S_IFDIR = 0040000; // directory
    const S_IFBLK = 0060000; // block special
    const S_IFREG = 0100000; // regular
    const S_IFLNK = 0120000; // symbolic link
    const S_IFSOCK = 0140000; // socket
    // setuid/setgid/sticky bits, the same as for chmod:
    const S_ISUID = 0004000; // set user id on execution
    const S_ISGID = 0002000; // set group id on execution
    const S_ISTXT = 0001000; // sticky bit
    // And of course, the other 12 bits are for the permissions, the same as for chmod:
    // When addding these up, you can also just write the permissions as a simgle octal number
    // ie. 0755. The leading 0 specifies octal notation.
    const S_IRWXU = 0000700; // RWX mask for owner
    const S_IRUSR = 0000400; // R for owner
    const S_IWUSR = 0000200; // W for owner
    const S_IXUSR = 0000100; // X for owner
    const S_IRWXG = 0000070; // RWX mask for group
    const S_IRGRP = 0000040; // R for group
    const S_IWGRP = 0000020; // W for group
    const S_IXGRP = 0000010; // X for group
    const S_IRWXO = 0000007; // RWX mask for other
    const S_IROTH = 0000004; // R for other
    const S_IWOTH = 0000002; // W for other
    const S_IXOTH = 0000001; // X for other
    const S_ISVTX = 0001000; // save swapped text even after use
    // Filetype, sticky and permissions are added up, and shifted 16 bits left BEFORE adding the DOS flags.
    // DOS file type flags, we really only use the S_DOS_D flag.
    const S_DOS_A = 0000040; // DOS flag for Archive
    const S_DOS_D = 0000020; // DOS flag for Directory
    const S_DOS_V = 0000010; // DOS flag for Volume
    const S_DOS_S = 0000004; // DOS flag for System
    const S_DOS_H = 0000002; // DOS flag for Hidden
    const S_DOS_R = 0000001; // DOS flag for Read Only

    private $zipMemoryThreshold = 1048576; // Autocreate tempfile if the zip data exceeds 1048576 bytes (1 MB)
    private $zipData = NULL;
    private $zipFile = NULL;
    private $zipComment = NULL;
    private $cdRec = array(); // central directory
    private $offset = 0;
    private $isFinalized = FALSE;
    private $addExtraField = TRUE;
    private $streamChunkSize = 65536;
    private $streamFilePath = NULL;
    private $streamTimestamp = NULL;
    private $streamFileComment = NULL;
    public $streamFile = NULL;
    public $streamData = NULL;
    private $streamFileLength = 0;
    private $streamExtFileAttr = null;

    /**
     * Constructor.
     *
     * @param boolean $useZipFile Write temp zip data to tempFile? Default FALSE
     */
    function __construct($useZipFile = FALSE)
    {
        if ($useZipFile)
        {
            $this->zipFile = tmpfile();
        }
        else
        {
            $this->zipData = "";
        }
    }

    function __destruct()
    {
        if (is_resource($this->zipFile))
        {
            fclose($this->zipFile);
        }
        $this->zipData = NULL;
    }

    /**
     * Extra fields on the Zip directory records are Unix time codes needed for compatibility on the default Mac zip archive tool.
     * These are enabled as default, as they do no harm elsewhere and only add 26 bytes per file added.
     *
     * @param bool $setExtraField TRUE (default) will enable adding of extra fields, anything else will disable it.
     */
    function setExtraField($setExtraField = TRUE)
    {
        $this->addExtraField = ($setExtraField === TRUE);
    }

    /**
     * Set Zip archive comment.
     *
     * @param string $newComment New comment. NULL to clear.
     * @return bool $success
     */
    public function setComment($newComment = NULL)
    {
        if ($this->isFinalized)
        {
            return FALSE;
        }
        $this->zipComment = $newComment;

        return TRUE;
    }

    /**
     * Set zip file to write zip data to.
     * This will cause all present and future data written to this class to be written to this file.
     * This can be used at any time, even after the Zip Archive have been finalized. Any previous file will be closed.
     * Warning: If the given file already exists, it will be overwritten.
     *
     * @param string $fileName
     * @return bool $success
     */
    public function setZipFile($fileName)
    {
        if (is_file($fileName))
        {
            unlink($fileName);
        }
        $fd = fopen($fileName, "x+b");
        if (is_resource($this->zipFile))
        {
            rewind($this->zipFile);
            while (!feof($this->zipFile))
            {
                fwrite($fd, fread($this->zipFile, $this->streamChunkSize));
            }

            fclose($this->zipFile);
        }
        else
        {
            fwrite($fd, $this->zipData);
            $this->zipData = NULL;
        }
        $this->zipFile = $fd;

        return TRUE;
    }

    /**
     * Add an empty directory entry to the zip archive.
     * Basically this is only used if an empty directory is added.
     *
     * @param string $directoryPath Directory Path and name to be added to the archive.
     * @param int    $timestamp     (Optional) Timestamp for the added directory, if omitted or set to 0, the current time will be used.
     * @param string $fileComment   (Optional) Comment to be added to the archive for this directory. To use fileComment, timestamp must be given.
     * @param int    $extFileAttr   (Optional) The external file reference, use generateExtAttr to generate this.
     * @return bool $success
     */
    public function addDirectory($directoryPath, $timestamp = 0, $fileComment = NULL, $extFileAttr = self::EXT_FILE_ATTR_DIR)
    {
        if ($this->isFinalized)
        {
            return FALSE;
        }
        $directoryPath = str_replace("\\", "/", $directoryPath);
        $directoryPath = rtrim($directoryPath, "/");

        if (strlen($directoryPath) > 0)
        {
            $this->buildZipEntry($directoryPath . '/', $fileComment, "\x00\x00", "\x00\x00", $timestamp, "\x00\x00\x00\x00", 0, 0, $extFileAttr);
            return TRUE;
        }
        return FALSE;
    }

    /**
     * Add a file to the archive at the specified location and file name.
     *
     * @param string $data        File data.
     * @param string $filePath    Filepath and name to be used in the archive.
     * @param int    $timestamp   (Optional) Timestamp for the added file, if omitted or set to 0, the current time will be used.
     * @param string $fileComment (Optional) Comment to be added to the archive for this file. To use fileComment, timestamp must be given.
     * @param bool   $compress    (Optional) Compress file, if set to FALSE the file will only be stored. Default TRUE.
     * @param int    $extFileAttr (Optional) The external file reference, use generateExtAttr to generate this.
     * @return bool $success
     */
    public function addFile($data, $filePath, $timestamp = 0, $fileComment = NULL, $compress = TRUE, $extFileAttr = self::EXT_FILE_ATTR_FILE)
    {
        if ($this->isFinalized)
        {
            return FALSE;
        }

        if (is_resource($data) && get_resource_type($data) == "stream")
        {
            $this->addLargeFile($data, $filePath, $timestamp, $fileComment, $extFileAttr);
            return FALSE;
        }

        $gzData = "";
        $gzType = "\x08\x00"; // Compression type 8 = deflate
        $gpFlags = "\x00\x00"; // General Purpose bit flags for compression type 8 it is: 0=Normal, 1=Maximum, 2=Fast, 3=super fast compression.
        $dataLength = strlen($data);
        $fileCRC32 = pack("V", crc32($data));

        if ($compress)
        {
            $gzTmp = gzcompress($data);
            $gzData = substr(substr($gzTmp, 0, strlen($gzTmp) - 4), 2); // gzcompress adds a 2 byte header and 4 byte CRC we can't use.
            // The 2 byte header does contain useful data, though in this case the 2 parameters we'd be interrested in will always be 8 for compression type, and 2 for General purpose flag.
            $gzLength = strlen($gzData);
        }
        else
        {
            $gzLength = $dataLength;
        }

        if ($gzLength >= $dataLength)
        {
            $gzLength = $dataLength;
            $gzData = $data;
            $gzType = "\x00\x00"; // Compression type 0 = stored
            $gpFlags = "\x00\x00"; // Compression type 0 = stored
        }

        if (!is_resource($this->zipFile) && ($this->offset + $gzLength) > $this->zipMemoryThreshold)
        {
            $this->zipflush();
        }

        $this->buildZipEntry($filePath, $fileComment, $gpFlags, $gzType, $timestamp, $fileCRC32, $gzLength, $dataLength, $extFileAttr);

        $this->zipwrite($gzData);

        return TRUE;
    }

    /**
     * Add the content to a directory.
     *
     * @author Adam Schmalhofer <Adam.Schmalhofer@gmx.de>
     * @author A. Grandt
     *
     * @param string $realPath       Path on the file system.
     * @param string $zipPath        Filepath and name to be used in the archive.
     * @param bool   $recursive      Add content recursively, default is TRUE.
     * @param bool   $followSymlinks Follow and add symbolic links, if they are accessible, default is TRUE.
     * @param array &$addedFiles     Reference to the added files, this is used to prevent duplicates, efault is an empty array.
     *                               If you start the function by parsing an array, the array will be populated with the realPath
     *                               and zipPath kay/value pairs added to the archive by the function.
     * @param bool   $overrideFilePermissions Force the use of the file/dir permissions set in the $extDirAttr
     * 							     and $extFileAttr parameters.
     * @param int    $extDirAttr     Permissions for directories.
     * @param int    $extFileAttr    Permissions for files.
     */
    public function addDirectoryContent($realPath, $zipPath, $recursive = TRUE, $followSymlinks = TRUE, &$addedFiles = array(), $overrideFilePermissions = FALSE, $extDirAttr = self::EXT_FILE_ATTR_DIR, $extFileAttr = self::EXT_FILE_ATTR_FILE)
    {
        if (file_exists($realPath) && !isset($addedFiles[realpath($realPath)]))
        {
            if (is_dir($realPath))
            {
                if ($overrideFilePermissions)
                {
                    $this->addDirectory($zipPath, 0, null, $extDirAttr);
                }
                else
                {
                    $this->addDirectory($zipPath, 0, null, self::getFileExtAttr($realPath));
                }
            }

            $addedFiles[realpath($realPath)] = $zipPath;

            $iter = new DirectoryIterator($realPath);
            foreach ($iter as $file)
            {
                if ($file->isDot())
                {
                    continue;
                }
                $newRealPath = $file->getPathname();
                $newZipPath = self::pathJoin($zipPath, $file->getFilename());

                if (file_exists($newRealPath) && ($followSymlinks === TRUE || !is_link($newRealPath)))
                {
                    if ($file->isFile())
                    {
                        $addedFiles[realpath($newRealPath)] = $newZipPath;
                        if ($overrideFilePermissions)
                        {
                            $this->addLargeFile($newRealPath, $newZipPath, 0, null, $extFileAttr);
                        }
                        else
                        {
                            $this->addLargeFile($newRealPath, $newZipPath, 0, null, self::getFileExtAttr($newRealPath));
                        }
                    }
                    else if ($recursive === TRUE)
                    {
                        $this->addDirectoryContent($newRealPath, $newZipPath, $recursive, $followSymlinks, $addedFiles, $overrideFilePermissions, $extDirAttr, $extFileAttr);
                    }
                    else
                    {
                        if ($overrideFilePermissions)
                        {
                            $this->addDirectory($zipPath, 0, null, $extDirAttr);
                        }
                        else
                        {
                            $this->addDirectory($zipPath, 0, null, self::getFileExtAttr($newRealPath));
                        }
                    }
                }
            }
        }
    }

    /**
     * Add a file to the archive at the specified location and file name.
     *
     * @param string $dataFile    File name/path.
     * @param string $filePath    Filepath and name to be used in the archive.
     * @param int    $timestamp   (Optional) Timestamp for the added file, if omitted or set to 0, the current time will be used.
     * @param string $fileComment (Optional) Comment to be added to the archive for this file. To use fileComment, timestamp must be given.
     * @param int    $extFileAttr (Optional) The external file reference, use generateExtAttr to generate this.
     * @return bool $success
     */
    public function addLargeFile($dataFile, $filePath, $timestamp = 0, $fileComment = NULL, $extFileAttr = self::EXT_FILE_ATTR_FILE)
    {
        if ($this->isFinalized)
        {
            return FALSE;
        }

        if (is_string($dataFile) && is_file($dataFile))
        {
            $this->processFile($dataFile, $filePath, $timestamp, $fileComment, $extFileAttr);
        }
        else if (is_resource($dataFile) && get_resource_type($dataFile) == "stream")
        {
            $fh = $dataFile;
            $this->openStream($filePath, $timestamp, $fileComment, $extFileAttr);

            while (!feof($fh))
            {
                $this->addStreamData(fread($fh, $this->streamChunkSize));
            }
            $this->closeStream($this->addExtraField);
        }
        return TRUE;
    }

    /**
     * Create a stream to be used for large entries.
     *
     * @param string $filePath    Filepath and name to be used in the archive.
     * @param int    $timestamp   (Optional) Timestamp for the added file, if omitted or set to 0, the current time will be used.
     * @param string $fileComment (Optional) Comment to be added to the archive for this file. To use fileComment, timestamp must be given.
     * @param int    $extFileAttr (Optional) The external file reference, use generateExtAttr to generate this.
     * @return bool $success
     */
    public function openStream($filePath, $timestamp = 0, $fileComment = null, $extFileAttr = self::EXT_FILE_ATTR_FILE)
    {
        if (!function_exists('sys_get_temp_dir'))
        {
            die("ERROR: Zip " . self::VERSION . " requires PHP version 5.2.1 or above if large files are used.");
        }

        if ($this->isFinalized)
        {
            return FALSE;
        }

        $this->zipflush();

        if (strlen($this->streamFilePath) > 0)
        {
            $this->closeStream();
        }

        $this->streamFile = tempnam(sys_get_temp_dir(), 'Zip');
        $this->streamData = fopen($this->streamFile, "wb");
        $this->streamFilePath = $filePath;
        $this->streamTimestamp = $timestamp;
        $this->streamFileComment = $fileComment;
        $this->streamFileLength = 0;
        $this->streamExtFileAttr = $extFileAttr;

        return TRUE;
    }

    /**
     * Add data to the open stream.
     *
     * @param string $data
     * @return mixed length in bytes added or FALSE if the archive is finalized or there are no open stream.
     */
    public function addStreamData($data)
    {
        if ($this->isFinalized || strlen($this->streamFilePath) == 0)
        {
            return FALSE;
        }

        $length = fwrite($this->streamData, $data, strlen($data));
        if ($length != strlen($data))
        {
            die("<p>Length mismatch</p>\n");
        }
        $this->streamFileLength += $length;

        return $length;
    }

    /**
     * Close the current stream.
     *
     * @return bool $success
     */
    public function closeStream()
    {
        if ($this->isFinalized || strlen($this->streamFilePath) == 0)
        {
            return FALSE;
        }

        fflush($this->streamData);
        fclose($this->streamData);

        $this->processFile($this->streamFile, $this->streamFilePath, $this->streamTimestamp, $this->streamFileComment, $this->streamExtFileAttr);

        $this->streamData = null;
        $this->streamFilePath = null;
        $this->streamTimestamp = null;
        $this->streamFileComment = null;
        $this->streamFileLength = 0;
        $this->streamExtFileAttr = null;

        // Windows is a little slow at times, so a millisecond later, we can unlink this.
        unlink($this->streamFile);

        $this->streamFile = null;

        return TRUE;
    }

    private function processFile($dataFile, $filePath, $timestamp = 0, $fileComment = null, $extFileAttr = self::EXT_FILE_ATTR_FILE)
    {
        if ($this->isFinalized)
        {
            return FALSE;
        }

        $tempzip = tempnam(sys_get_temp_dir(), 'ZipStream');

        $zip = new ZipArchive;
        if ($zip->open($tempzip) === TRUE)
        {
            $zip->addFile($dataFile, 'file');
            $zip->close();
        }

        $file_handle = fopen($tempzip, "rb");
        $stats = fstat($file_handle);
        $eof = $stats['size'] - 72;

        fseek($file_handle, 6);

        $gpFlags = fread($file_handle, 2);
        $gzType = fread($file_handle, 2);
        fread($file_handle, 4);
        $fileCRC32 = fread($file_handle, 4);
        $v = unpack("Vval", fread($file_handle, 4));
        $gzLength = $v['val'];
        $v = unpack("Vval", fread($file_handle, 4));
        $dataLength = $v['val'];

        $this->buildZipEntry($filePath, $fileComment, $gpFlags, $gzType, $timestamp, $fileCRC32, $gzLength, $dataLength, $extFileAttr);

        fseek($file_handle, 34);
        $pos = 34;

        while (!feof($file_handle) && $pos < $eof)
        {
            $datalen = $this->streamChunkSize;
            if ($pos + $this->streamChunkSize > $eof)
            {
                $datalen = $eof - $pos;
            }
            $data = fread($file_handle, $datalen);
            $pos += $datalen;

            $this->zipwrite($data);
        }

        fclose($file_handle);

        unlink($tempzip);
    }

    /**
     * Close the archive.
     * A closed archive can no longer have new files added to it.
     *
     * @return bool $success
     */
    public function finalize()
    {
        if (!$this->isFinalized)
        {
            if (strlen($this->streamFilePath) > 0)
            {
                $this->closeStream();
            }
            $cd = implode("", $this->cdRec);

            $cdRecSize = pack("v", sizeof($this->cdRec));
            $cdRec = $cd . self::ZIP_END_OF_CENTRAL_DIRECTORY
                    . $cdRecSize . $cdRecSize
                    . pack("VV", strlen($cd), $this->offset);
            if (!empty($this->zipComment))
            {
                $cdRec .= pack("v", strlen($this->zipComment)) . $this->zipComment;
            }
            else
            {
                $cdRec .= "\x00\x00";
            }

            $this->zipwrite($cdRec);

            $this->isFinalized = TRUE;
            $this->cdRec = NULL;

            return TRUE;
        }
        return FALSE;
    }

    /**
     * Get the handle ressource for the archive zip file.
     * If the zip haven't been finalized yet, this will cause it to become finalized
     *
     * @return zip file handle
     */
    public function getZipFile()
    {
        if (!$this->isFinalized)
        {
            $this->finalize();
        }

        $this->zipflush();

        rewind($this->zipFile);

        return $this->zipFile;
    }

    /**
     * Get the zip file contents
     * If the zip haven't been finalized yet, this will cause it to become finalized
     *
     * @return zip data
     */
    public function getZipData()
    {
        if (!$this->isFinalized)
        {
            $this->finalize();
        }
        if (!is_resource($this->zipFile))
        {
            return $this->zipData;
        }
        else
        {
            rewind($this->zipFile);
            $filestat = fstat($this->zipFile);
            return fread($this->zipFile, $filestat['size']);
        }
    }

    /**
     * Send the archive as a zip download
     *
     * @param String $fileName The name of the Zip archive, in ISO-8859-1 (or ASCII) encoding, ie. "archive.zip". Optional, defaults to NULL, which means that no ISO-8859-1 encoded file name will be specified.
     * @param String $contentType Content mime type. Optional, defaults to "application/zip".
     * @param String $utf8FileName The name of the Zip archive, in UTF-8 encoding. Optional, defaults to NULL, which means that no UTF-8 encoded file name will be specified.
     * @param bool $inline Use Content-Disposition with "inline" instead of "attached". Optional, defaults to FALSE.
     * @return bool $success
     */
    function sendZip($fileName = null, $contentType = "application/zip", $utf8FileName = null, $inline = false)
    {
        if (!$this->isFinalized)
        {
            $this->finalize();
        }

        $headerFile = null;
        $headerLine = null;
        if (!headers_sent($headerFile, $headerLine) or die("<p><strong>Error:</strong> Unable to send file $fileName. HTML Headers have already been sent from <strong>$headerFile</strong> in line <strong>$headerLine</strong></p>"))
        {
            if ((ob_get_contents() === FALSE || ob_get_contents() == '') or die("\n<p><strong>Error:</strong> Unable to send file <strong>$fileName</strong>. Output buffer contains the following text (typically warnings or errors):<br>" . htmlentities(ob_get_contents()) . "</p>"))
            {
                if (ini_get('zlib.output_compression'))
                {
                    ini_set('zlib.output_compression', 'Off');
                }

                header("Pragma: public");
                header("Last-Modified: " . gmdate("D, d M Y H:i:s T"));
                header("Expires: 0");
                header("Accept-Ranges: bytes");
                header("Connection: close");
                header("Content-Type: " . $contentType);
                $cd = "Content-Disposition: ";
                if ($inline)
                {
                    $cd .= "inline";
                }
                else
                {
                    $cd .= "attached";
                }
                if ($fileName)
                {
                    $cd .= '; filename="' . $fileName . '"';
                }
                if ($utf8FileName)
                {
                    $cd .= "; filename*=UTF-8''" . rawurlencode($utf8FileName);
                }
                header($cd);
                header("Content-Length: " . $this->getArchiveSize());

                if (!is_resource($this->zipFile))
                {
                    echo $this->zipData;
                }
                else
                {
                    rewind($this->zipFile);

                    while (!feof($this->zipFile))
                    {
                        echo fread($this->zipFile, $this->streamChunkSize);
                    }
                }
            }
            return TRUE;
        }
        return FALSE;
    }

    /**
     * Return the current size of the archive
     *
     * @return $size Size of the archive
     */
    public function getArchiveSize()
    {
        if (!is_resource($this->zipFile))
        {
            return strlen($this->zipData);
        }
        $filestat = fstat($this->zipFile);

        return $filestat['size'];
    }

    /**
     * Calculate the 2 byte dostime used in the zip entries.
     *
     * @param int $timestamp
     * @return 2-byte encoded DOS Date
     */
    private function getDosTime($timestamp = 0)
    {
        $timestamp = (int) $timestamp;
        $oldTZ = @date_default_timezone_get();
        date_default_timezone_set('UTC');
        $date = ($timestamp == 0 ? getdate() : getdate($timestamp));
        date_default_timezone_set($oldTZ);
        if ($date["year"] >= 1980)
        {
            return pack("V", (($date["mday"] + ($date["mon"] << 5) + (($date["year"] - 1980) << 9)) << 16) |
                            (($date["seconds"] >> 1) + ($date["minutes"] << 5) + ($date["hours"] << 11)));
        }
        return "\x00\x00\x00\x00";
    }

    /**
     * Build the Zip file structures
     *
     * @param string $filePath
     * @param string $fileComment
     * @param string $gpFlags
     * @param string $gzType
     * @param int    $timestamp
     * @param string $fileCRC32
     * @param int    $gzLength
     * @param int    $dataLength
     * @param int    $extFileAttr Use self::EXT_FILE_ATTR_FILE for files, self::EXT_FILE_ATTR_DIR for Directories.
     */
    private function buildZipEntry($filePath, $fileComment, $gpFlags, $gzType, $timestamp, $fileCRC32, $gzLength, $dataLength, $extFileAttr)
    {
        $filePath = str_replace("\\", "/", $filePath);
        $fileCommentLength = (empty($fileComment) ? 0 : strlen($fileComment));
        $timestamp = (int) $timestamp;
        $timestamp = ($timestamp == 0 ? time() : $timestamp);

        $dosTime = $this->getDosTime($timestamp);
        $tsPack = pack("V", $timestamp);

        if (!isset($gpFlags) || strlen($gpFlags) != 2)
        {
            $gpFlags = "\x00\x00";
        }

        $isFileUTF8 = true;
        $isCommentUTF8 = !empty($fileComment) && true;

        $localExtraField = "";
        $centralExtraField = "";

        if ($this->addExtraField)
        {
            $localExtraField .= "\x55\x54\x09\x00\x03" . $tsPack . $tsPack . Zip::EXTRA_FIELD_NEW_UNIX_GUID;
            $centralExtraField .= "\x55\x54\x05\x00\x03" . $tsPack . Zip::EXTRA_FIELD_NEW_UNIX_GUID;
        }

        if ($isFileUTF8 || $isCommentUTF8)
        {
            $flag = 0;
            $gpFlagsV = unpack("vflags", $gpFlags);
            if (isset($gpFlagsV['flags']))
            {
                $flag = $gpFlagsV['flags'];
            }
            $gpFlags = pack("v", $flag | (1 << 11));

            if ($isFileUTF8)
            {
                $utfPathExtraField = "\x75\x70"
                        . pack("v", (5 + strlen($filePath)))
                        . "\x01"
                        . pack("V", crc32($filePath))
                        . $filePath;

                $localExtraField .= $utfPathExtraField;
                $centralExtraField .= $utfPathExtraField;
            }
            if ($isCommentUTF8)
            {
                $centralExtraField .= "\x75\x63" // utf8 encoded file comment extra field
                        . pack("v", (5 + strlen($fileComment)))
                        . "\x01"
                        . pack("V", crc32($fileComment))
                        . $fileComment;
            }
        }

        $header = $gpFlags . $gzType . $dosTime . $fileCRC32
                . pack("VVv", $gzLength, $dataLength, strlen($filePath)); // File name length

        $zipEntry = self::ZIP_LOCAL_FILE_HEADER
                . self::ATTR_VERSION_TO_EXTRACT
                . $header
                . pack("v", strlen($localExtraField)) // Extra field length
                . $filePath // FileName
                . $localExtraField; // Extra fields

        $this->zipwrite($zipEntry);

        $cdEntry = self::ZIP_CENTRAL_FILE_HEADER
                . self::ATTR_MADE_BY_VERSION
                . ($dataLength === 0 ? "\x0A\x00" : self::ATTR_VERSION_TO_EXTRACT)
                . $header
                . pack("v", strlen($centralExtraField)) // Extra field length
                . pack("v", $fileCommentLength) // File comment length
                . "\x00\x00" // Disk number start
                . "\x00\x00" // internal file attributes
                . pack("V", $extFileAttr) // External file attributes
                . pack("V", $this->offset) // Relative offset of local header
                . $filePath // FileName
                . $centralExtraField; // Extra fields

        if (!empty($fileComment))
        {
            $cdEntry .= $fileComment; // Comment
        }

        $this->cdRec[] = $cdEntry;
        $this->offset += strlen($zipEntry) + $gzLength;
    }

    private function zipwrite($data)
    {
        if (!is_resource($this->zipFile))
        {
            $this->zipData .= $data;
        }
        else
        {
            fwrite($this->zipFile, $data);
            fflush($this->zipFile);
        }
    }

    private function zipflush()
    {
        if (!is_resource($this->zipFile))
        {
            $this->zipFile = tmpfile();
            fwrite($this->zipFile, $this->zipData);
            $this->zipData = NULL;
        }
    }

    /**
     * Join $file to $dir path, and clean up any excess slashes.
     *
     * @param string $dir
     * @param string $file
     */
    public static function pathJoin($dir, $file)
    {
        if (empty($dir) || empty($file))
        {
            return self::getRelativePath($dir . $file);
        }
        return self::getRelativePath($dir . '/' . $file);
    }

    /**
     * Clean up a path, removing any unnecessary elements such as /./, // or redundant ../ segments.
     * If the path starts with a "/", it is deemed an absolute path and any /../ in the beginning is stripped off.
     * The returned path will not end in a "/".
     *
     * Sometimes, when a path is generated from multiple fragments, 
     *  you can get something like "../data/html/../images/image.jpeg"
     * This will normalize that example path to "../data/images/image.jpeg"
     *
     * @param string $path The path to clean up
     * @return string the clean path
     */
    public static function getRelativePath($path)
    {
        $path = preg_replace("#/+\.?/+#", "/", str_replace("\\", "/", $path));
        $dirs = explode("/", rtrim(preg_replace('#^(?:\./)+#', '', $path), '/'));

        $offset = 0;
        $sub = 0;
        $subOffset = 0;
        $root = "";

        if (empty($dirs[0]))
        {
            $root = "/";
            $dirs = array_splice($dirs, 1);
        }
        else if (preg_match("#[A-Za-z]:#", $dirs[0]))
        {
            $root = strtoupper($dirs[0]) . "/";
            $dirs = array_splice($dirs, 1);
        }

        $newDirs = array();
        foreach ($dirs as $dir)
        {
            if ($dir !== "..")
            {
                $subOffset--;
                $newDirs[++$offset] = $dir;
            }
            else
            {
                $subOffset++;
                if (--$offset < 0)
                {
                    $offset = 0;
                    if ($subOffset > $sub)
                    {
                        $sub++;
                    }
                }
            }
        }

        if (empty($root))
        {
            $root = str_repeat("../", $sub);
        }
        return $root . implode("/", array_slice($newDirs, 0, $offset));
    }

    /**
     * Create the file permissions for a file or directory, for use in the extFileAttr parameters.
     *
     * @param int   $owner Unix permisions for owner (octal from 00 to 07)
     * @param int   $group Unix permisions for group (octal from 00 to 07)
     * @param int   $other Unix permisions for others (octal from 00 to 07)
     * @param bool  $isFile
     * @return EXTRERNAL_REF field.
     */
    public static function generateExtAttr($owner = 07, $group = 05, $other = 05, $isFile = true)
    {
        $fp = $isFile ? self::S_IFREG : self::S_IFDIR;
        $fp |= (($owner & 07) << 6) | (($group & 07) << 3) | ($other & 07);

        return ($fp << 16) | ($isFile ? self::S_DOS_A : self::S_DOS_D);
    }

    /**
     * Get the file permissions for a file or directory, for use in the extFileAttr parameters.
     *
     * @param string $filename
     * @return external ref field, or FALSE if the file is not found.
     */
    public static function getFileExtAttr($filename)
    {
        if (file_exists($filename))
        {
            $fp = fileperms($filename) << 16;
            return $fp | (is_dir($filename) ? self::S_DOS_D : self::S_DOS_A);
        }
        return FALSE;
    }

}

?>
