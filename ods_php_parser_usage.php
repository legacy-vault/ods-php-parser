<?php

    // ods_php_parser_usage.php
    
    // Shows an Example how to use ODS PHP Parser.

include("ods_php_parser.php");

/////////////////////////////////////////////////
/*
// Disable Buffering
ini_set('output_buffering', 'off');
ini_set('zlib.output_compression', false);
while (@ob_end_flush());
ini_set('implicit_flush', true);
ob_implicit_flush(true);
header("Content-type: text/plain");
header('Cache-Control: no-cache'); 
*/
/////////////////////////////////////////////////

$nl = "<br>\r\n";
$fn = $_SERVER['DOCUMENT_ROOT'] . '/test.ods';  // FileName
$my_ods;
$result;
$k;
$v;
$i; $i_max;
$j; $j_max;
$val;
$range;
$bg;


echo "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>\r\n";

//---------------------------------------

// Open
$my_ods = new ODS();
$result = $my_ods->Open($fn);
if ($result == FALSE)
{
    echo "Error opening." . $nl; //
    return;
}

//---------------------------------------

// Parse
$my_ods->Parse();

//---------------------------------------

// Close
$result = $my_ods->Close();
if ($result == FALSE)
{
    echo "Error closing." . $nl; //
    return;
}

//---------------------------------------

// Sheets Count
echo "Sheets in File: " . $my_ods->SheetCount . $nl;

//---------------------------------------

// Columns Count
echo "Columns in Sheets: ";
foreach ($my_ods->Sheets as $k => $v)
{
    echo "[$k:" . $v[ ODS::COL_COUNT ] . ']';
}
echo $nl;

//---------------------------------------

// Rows Count
echo "Rows in Sheets: ";
foreach ($my_ods->Sheets as $k => $v)
{
    echo "[$k:" . $v[ ODS::ROW_COUNT ] . ']';
}
echo $nl;

//---------------------------------------

// Show Tables (hidden Cells are shown as if no span enabled)
echo $nl;
for ($i = 1; $i <= $my_ods->SheetCount; $i++)
{
    echo "<b>Sheet #$i</b> (" . $my_ods->Sheets[$i][ODS::ATTRIBUTES][ODS::ATTRIB_TABLE_NAME] . ")$nl";
    
    // UsedRange
    echo "Used Cols: from " . $my_ods->Sheets[$i][ODS::FIRST_USED_COL] . " to " . $my_ods->Sheets[$i][ODS::LAST_USED_COL] . '.' . $nl;
    echo "Used Rows: from " . $my_ods->Sheets[$i][ODS::FIRST_USED_ROW] . " to " . $my_ods->Sheets[$i][ODS::LAST_USED_ROW] . '.' . $nl;
    
    echo "<table border='1' cellpadding='1' cellspacing='0'>";
    for ($j = $my_ods->Sheets[$i][ODS::FIRST_USED_ROW]; $j <= $my_ods->Sheets[$i][ODS::LAST_USED_ROW]; $j++)
    {
        echo "<tr>";
        for ($k = $my_ods->Sheets[$i][ODS::FIRST_USED_COL]; $k <= $my_ods->Sheets[$i][ODS::LAST_USED_COL]; $k++)
        {
            $val = '';
            if ( isset( $my_ods->Sheets[$i][ODS::CELLS][$j][$k][ODS::VALUE] ) )
            {
                $val = $my_ods->Sheets[$i][ODS::CELLS][$j][$k][ODS::VALUE];
            }
            if ($val == '')
            {
                $val = '&nbsp;';
            }
            else
            {
                $val = htmlentities($val, ENT_SUBSTITUTE | ENT_HTML5, 'UTF-8');
            }
            
            echo "<td>$val</td>\r\n";
        }
        echo "</tr>\r\n";
    }
    echo "</table>" . $nl;
}

//---------------------------------------

// Get a Cell
echo "<b>Get Cell Examples.</b>$nl";
$val = $my_ods->GetCell(1, 1, 1);
echo "Cell on Sheet [1], at Row [1], at Column [1] is: $val" . $nl;
$val = $my_ods->GetCell(2, 8, 4);
echo "Cell on Sheet [2], at Row [8], at Column [4] is: $val" . $nl;
$val = $my_ods->GetCell(5, 16, 33);
if ( !is_null($val) )
{
    echo "Cell on Sheet [5], at Row [16], at Column [33] is: $val" . $nl;
}
else
{
    echo "Cell on Sheet [5], at Row [16], at Column [33] is not set!" . $nl;
}
echo $nl;

//---------------------------------------

// Get Used Range of a Sheet and show it
echo "<b>Used Range of Sheet # 2.</b>$nl";
echo "(Set Cells are in light green Color, non-set Cells are grey)$nl";
$range = $my_ods->GetUsedRange(2);
if ( is_null($range) )
{
    echo "The Range is not set!$nl";
}
else
{
    $i_max = $range[ODS::ROW_COUNT];
    $j_max = $range[ODS::COL_COUNT];
    
    echo "<table border='1' cellpadding='1' cellspacing='0'>";
    for ($i = 1; $i <= $i_max; $i++)
    {
        echo "<tr>";
        for ($j = 1; $j <= $j_max; $j++)
        {
            $val = $range[$i][$j];
            if ( is_null( $val ) )
            {
                $val = '&nbsp;';
                $bg = 'grey';
            }
            else
            {
                $bg = 'lightgreen';
            }
            echo "<td bgcolor='$bg'>$val</td>\r\n";
        }
        echo "</tr>\r\n";
    }
    echo "</table>" . $nl;
}

//---------------------------------------

echo "</body></html>";

?>
