<?php

    // ods_php_parser.php
    
    // Simple ODS PHP Parser.
    
    // Version: 0.1.
    // Date: 2017-05-09.
    // Author: McArcher.

class ODS
{
    // This Class implements simple Parsing of Data from zipped ODS Files.
    // It can open the ODS File, parse (read) Data and close the ODS File.
    // Not all the Attributes and Tags mentioned in the Code are used yet.
    // For more Understanding, also see 'ods_php_parser_usage.php'.

    // Constants
    const UNZIPPED_SUBFOLDER = 'unz';
    
    // Attribute Names
    const ATTRIB_REPEAT_H = 'table:number-columns-repeated';
    const ATTRIB_REPEAT_V = 'table:number-rows-repeated';
    const ATTRIB_SPAN_H = 'table:number-columns-spanned';
    const ATTRIB_SPAN_V = 'table:number-rows-spanned';
    const ATTRIB_TABLE_NAME = 'table:name';
    
    // Tag Names
    const TAG_TABLE = 'table:table';
    const TAG_TITLE = 'table:title';
    const TAG_DESCRIPTION = 'table:desc';
    
    const TAG_ROW = 'table:table-row';
    const TAG_ROWS = 'table:table-rows';
    const TAG_ROW_GROUP = 'table:table-row-group';
    
    const TAG_CELL = 'table:table-cell';
    const TAG_CELL_COVERED = 'table:covered-table-cell';
    
    const TAG_COLUMN = 'table:table-column';
    const TAG_COLUMNS = 'table:table-columns';
    const TAG_COLUMN_GROUP = 'table:table-column-group';
    
    const TAG_HEADER_COLUMNS = 'table:table-header-columns';
    const TAG_HEADER_ROWS = 'table:table-header-rows';
    
    // Names of Fields in 'Sheets' Array
    const COL_COUNT = 'cc';
    const ROW_COUNT = 'rc';
    const COLS = 'cols';
    const ROWS = 'rows';
    const CELLS = 'cells';
    const FIRST_USED_COL = 'fuc';   // When you select the whole Row, LibreOffice (v3.3.3+) fills all 1024 Cells in a Row!
    const LAST_USED_COL = 'luc';    // When you select the whole Column, LibreOffice (v3.3.3+) fills all 1 048 576 Cells in a Column!
    const FIRST_USED_ROW = 'fur';   // That is why we need a Mechanism to ignore empty Cells. 
    const LAST_USED_ROW = 'lur';    // Otherwise it would take Hours and Days to manipulate all of them :D
    const ATTRIBUTES = 'atr';
    const VALUE = 'val';
    
    // Public Variables

    // Arrays
    public $Sheets;

    // Indicators
    public $WorkingFolder;
    public $WorkingFile;
    public $SheetCount;

    // Internal Variables
    protected $currentSheet;    // starts from 1
    protected $currentCol;      // starts from 1
    protected $currentRow;      // starts from 1
    protected $currentCell;     // starts from 1
    protected $repeatCol;
    protected $repeatRow;
    protected $workWithCell;

    //---------------------------------------

    public function __construct()
    {
        // Arrays
        $this->Sheets = array();
        
        // Indicators
        $this->WorkingFolder = '';
        $this->WorkingFile = '';
        $this->SheetCount = 0;
        
        // Internal Parameters
        $this->currentSheet = 0;
        $this->currentCol = 0;
        $this->currentRow = 0;
        $this->currentCell = 0;
        $this->repeatCol = 0;
        $this->repeatRow = 0;
        $this->workWithCell = FALSE;
    }

    //---------------------------------------

    public function Open($file_input)
    {
        // Opens the File and unpacks it to a temporary Directory.
        // If Errors occur, it returns FALSE, otherwise returns TRUE.
        
        $dir_tmp;
        $dir_new;
        $attempts;
        $result;
        $dir_unzipped;
        $path_zip;
        $cmd;
        $obj;
        $path_content_xml;
        $xml_contents;
        $output;
        $exit_code;
        
        $dir_tmp = sys_get_temp_dir();
        
        // Create random Name for Directory in System's temporary Directory
        $dir_new = tempnam($dir_tmp, '');
        $attempts = 1;
        while (!$dir_new)
        {
            if ($attempts > 100)
            {
                echo "Error! Cannot create temporary folder inside $dir_tmp! Too many attempts used."; //
                return FALSE;
            }
            $dir_new = tempnam($dir_tmp, '');
            $attempts++;
        }
        
        // While 'tempnam' creates a File, we need to change it to Directory
        $result = unlink($dir_new);
        
        if (!$result)
        {
            echo "Error deleting temporary folder $dir_new inside $dir_tmp! 'unlink' failed."; //
            return FALSE;
        }        
        
        $result = mkdir($dir_new);
        
        if (!$result)
        {
            echo "Error! Cannot create temporary folder $dir_new inside $dir_tmp! 'mkdir' failed."; //
            return FALSE;
        }
        
        $dir_unzipped = $dir_new . '/' . self::UNZIPPED_SUBFOLDER;
        $result = mkdir($dir_unzipped);
        
        if (!$result)
        {
            echo "Error! Cannot create temporary folder $dir_unzipped inside $dir_tmp! 'mkdir' failed."; //
            return FALSE;
        }
        
        // Save Paths into Object
        $this->WorkingFolder = $dir_new;            // ~'/tmp/rndxyz'
        $this->WorkingFile = basename($file_input);  // ~'my.ods'
        
        // Copy zipped File
        $path_zip = $this->WorkingFolder . '/' . $this->WorkingFile; 
        copy($file_input, $path_zip);
        
        // Unpack zipped File
        $cmd = 'unzip ' . escapeshellarg($path_zip) . ' -d ' . escapeshellarg($dir_unzipped);        
        exec($cmd, $output, $exit_code);
        if ($exit_code != 0)
        {
            echo "Error! Cannot unpack archive $path_zip! 'unzip' failed."; //
            return FALSE;
        }
                
        return TRUE;
    }
    
    //---------------------------------------
    
    public function Close()
    {
        $cmd;
        $output;
        $exit_code;
    
        $cmd = 'rm -r ' . escapeshellarg($this->WorkingFolder);        
        exec($cmd, $output, $exit_code);
        if ($exit_code != 0)
        {
            echo 'Error! Cannot delete directory ' . $this->WorkingFolder . "! 'rm' failed."; //
            return FALSE;
        }
        
        return TRUE;
    }
    
    //---------------------------------------

    public function Parse()
    {
        $dir_unzipped;
        $path_content_xml;
        $xml_contents;
        $xml_parser;
        
        $dir_unzipped = $this->WorkingFolder . '/' . self::UNZIPPED_SUBFOLDER;
        $path_content_xml = $dir_unzipped . '/' . 'content.xml';
        $xml_contents = file_get_contents($path_content_xml);
        
        // Create Parser
        $xml_parser = xml_parser_create(); 
        
        // Bind with Object
        xml_set_object ($xml_parser, $this);
        
        // Set Handlers
        xml_set_element_handler($xml_parser, "startElement", "endElement");
        xml_set_character_data_handler($xml_parser, "characterData");
        xml_parser_set_option($xml_parser, XML_OPTION_CASE_FOLDING, FALSE); // Case of Attribute's Names is not "uppered"
        
        // Do Parsing
        xml_parse($xml_parser, $xml_contents, strlen($xml_contents));
        
        // Clean up
        xml_parser_free($xml_parser);
        unset($xml_parser);
    }

    //---------------------------------------

    protected function startElement($parser, $tagName, $attribs)
    {
        // Handler for opening Tags.
        
        $n;
        $m;
        $x; $x_max;
        $y; $y_max;
        
        $tagName = strtolower($tagName);
        
        // Examine Tags
        // N.B.: '$this' is an Object of 'ODS' Class, not the XML-Parser.
        if($tagName == self::TAG_TABLE)
        {
            $this->SheetCount++;
            $this->currentSheet = $this->SheetCount;   
            $this->currentRow = 1;
            
            // Fill Object
            $this->Sheets[$this->currentSheet][self::ATTRIBUTES] = $attribs;
        }
        elseif ($tagName == self::TAG_COLUMN)
        {
            $n = 1;            
            if ( isset( $attribs[self::ATTRIB_REPEAT_H] ) )
            {
                $n = intval( $attribs[self::ATTRIB_REPEAT_H] );
            }
            $this->currentCol += $n;    // These Tags go above all others, so we 
                                        // do not need to save Cursor to the first Column
            
            // Fill Object
            $this->Sheets[$this->currentSheet][self::COL_COUNT] = $this->currentCol;
            $this->Sheets[$this->currentSheet][self::COLS][$this->currentCol][self::ATTRIBUTES] = $attribs;
        }
        elseif ($tagName == self::TAG_ROW)
        {            
            // repeatRow
            $m = 1;
            if ( isset( $attribs[self::ATTRIB_REPEAT_V] ) )
            {
                $m = intval( $attribs[self::ATTRIB_REPEAT_V] );
            }
            $this->repeatRow = $m;  // Provide Parameter for other Handlers
            
            $this->currentCol = 1;
            
            // Fill Object
            $this->Sheets[$this->currentSheet][self::ROW_COUNT] = $this->currentRow + ($m - 1);
            $this->Sheets[$this->currentSheet][self::ROWS][$this->currentRow][self::ATTRIBUTES] = $attribs;
        }
        elseif ( ($tagName == self::TAG_CELL) || ($tagName == self::TAG_CELL_COVERED) )
        {
            // workWithCell
            $this->workWithCell = TRUE;
            
            // repeatCol
            $n = 1;            
            if ( isset( $attribs[self::ATTRIB_REPEAT_H] ) )
            {
                $n = intval( $attribs[self::ATTRIB_REPEAT_H] );
            }            
            $this->repeatCol = $n;  // Provide Parameter for other Handlers
            
            // Fill Object
            $this->Sheets[$this->currentSheet][self::CELLS][$this->currentRow][$this->currentCol][self::ATTRIBUTES] = $attribs;
        }
    }

    //---------------------------------------

    protected function endElement($parser, $tagName)
    {
        // Handler for closing Tags.
        
        $tagName = strtolower($tagName);
        
        if($tagName == self::TAG_TABLE)
        {
            $this->currentSheet = 0; // Between Sheets is nothing
            $this->currentCol = 0;
            $this->currentRow = 0;
        }
        elseif ($tagName == self::TAG_ROW)
        {
            // Cursor
            $this->currentRow += $this->repeatRow;
        }
        elseif ( ($tagName == self::TAG_CELL) || ($tagName == self::TAG_CELL_COVERED) )
        {
            // workWithCell
            $this->workWithCell = FALSE;
            
            // Cursor
            $this->currentCol += $this->repeatCol;
        }
    }

    //---------------------------------------

    protected function characterData($parser, $data)
    {
        // Handler for Data between Tags.
        
        $x; $x_max;
        $y; $y_max;
        
        // Empty Tags are ignored to save Memory
        if ( strlen($data) == 0 )
        {
            return;
        }
        
        // Cells
        if ( $this->workWithCell )
        {
            // Restore saved Parameters
            $y = $this->currentRow;
            $y_max = $y + $this->repeatRow - 1;
            $x = $this->currentCol;
            $x_max = $x + $this->repeatCol - 1;
            
            // UsedRange
            
            // FIRST_USED_COL
            if ( isset( $this->Sheets[$this->currentSheet][self::FIRST_USED_COL] ) )
            {
                if ( $x < $this->Sheets[$this->currentSheet][self::FIRST_USED_COL] )
                {
                    $this->Sheets[$this->currentSheet][self::FIRST_USED_COL] = $x;
                }
            }
            else
            {
                $this->Sheets[$this->currentSheet][self::FIRST_USED_COL] = $x;
            }
            
            // LAST_USED_COL
            if ( isset( $this->Sheets[$this->currentSheet][self::LAST_USED_COL] ) )
            {
                if ( $this->Sheets[$this->currentSheet][self::LAST_USED_COL] < $x_max )
                {
                    $this->Sheets[$this->currentSheet][self::LAST_USED_COL] = $x_max;
                }
            }
            else
            {
                $this->Sheets[$this->currentSheet][self::LAST_USED_COL] = $x_max;
            }
            
            // FIRST_USED_ROW
            if ( isset( $this->Sheets[$this->currentSheet][self::FIRST_USED_ROW] ) )
            {
                if ( $y < $this->Sheets[$this->currentSheet][self::FIRST_USED_ROW] )
                {
                    $this->Sheets[$this->currentSheet][self::FIRST_USED_ROW] = $y;
                }
            }
            else
            {
                $this->Sheets[$this->currentSheet][self::FIRST_USED_ROW] = $y;
            }
            
            // LAST_USED_ROW
            if ( isset( $this->Sheets[$this->currentSheet][self::LAST_USED_ROW] ) )
            {
                if ( $this->Sheets[$this->currentSheet][self::LAST_USED_ROW] < $y_max )
                {
                    $this->Sheets[$this->currentSheet][self::LAST_USED_ROW] = $y_max;
                }
            }
            else
            {
                $this->Sheets[$this->currentSheet][self::LAST_USED_ROW] = $y_max;
            }
            
            // Fill Object
            while ($y <= $y_max)
            {
                $x = $this->currentCol;
                while ($x <= $x_max)
                {
                    // Unfortunately, this built-in-PHP XML Parser gives Data in Portions, 
                    // not as a whole Piece! So, we must join Pieces into one.
                    if ( isset($this->Sheets[$this->currentSheet][self::CELLS][$y][$x][self::VALUE]) )
                    {
                        $this->Sheets[$this->currentSheet][self::CELLS][$y][$x][self::VALUE] .= $data;
                    }
                    else
                    {
                        $this->Sheets[$this->currentSheet][self::CELLS][$y][$x][self::VALUE] = $data;
                    }                    
                    $x++;
                }                
                $y++;
            }
        }
    }

    //---------------------------------------

} // End of Class 'ODS'

//---------------------------------------

?>
