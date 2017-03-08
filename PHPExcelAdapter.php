<?php
// Prevent redeclaring this class - causes a fatal error.
if(!class_exists('PHPExcelAdapter'))
{
require_once "PHPExcel.php";

class PHPExcelAdapter
{
    /**
    * The standard phpExcel object from PHPExcel.php<br>
    * <i>This object can be used to perform any operations found in the standard PHPExcel class</i>
     * 
     * @link https://github.com/PHPOffice/PHPExcel/wiki/User%20Documentation
     * @link http://clock.co.uk/tech-blogs/phpexcel-cheatsheet
    */
    public $phpExcel;
    
    public function __construct()
    {
        $this->phpExcel = new PHPExcel();
    }
    
    /***************************************** Standard functions ********************************/
    
    /**
    * Sets the document's internal properties<br>
    * -Creator, Last Modified By, Keywords, Category = $creator<br>
    * -Title, Subject, Description = $title
    * 
    * @param string $creator <br>
    * @param string $title
    * 
    * @return void
    */
    public function setSpreadsheetProperties($creator,$title)
    {
        $this->phpExcel->getProperties()->setCreator($creator)
                                ->setLastModifiedBy($creator)
                                ->setTitle(strtoupper($title))
                                ->setSubject($title)
                                ->setDescription(strtoupper($title))
                                ->setKeywords($creator)
                                ->setCategory($creator);
    }
    
    /**
     * Draws an image at the specified cell (Mainly used for logos)<br>
     * 
     * @param string $cell (A1)<br>
     * @param string $name (Image name)<br>
     * @param string $description (Image description)<br>
     * @param string $image_path (Location of image)<br>
     * @param int $height = 50<br>
     * @param int $offsetX = 10<br>
     * @param int $offsetY = 10<br>
     * 
     * 
     * @return void
     */
    public function drawImage($cell, $name, $description, $image_path, $height=50, $offsetX=10, $offsetY=10)
    {
        $objDrawing = new PHPExcel_Worksheet_Drawing();
        $objDrawing->setName($name);
        $objDrawing->setDescription($description);
        $objDrawing->setPath($image_path);
        $objDrawing->setCoordinates($cell);
        $objDrawing->setHeight($height);
        $objDrawing->setOffsetX($offsetX);
        $objDrawing->setOffsetY($offsetY);
        $objDrawing->setWorksheet($this->phpExcel->getActiveSheet());
    }
    
    /**
     * Sets the sheet to be used<br>
     * <i>Would be nice to implement it to dynamically change to the sheet by sheet_name</i>
     * 
     * @param int $sheet_index = 0<br>
     * 
     * @return void
     */
    public function setActiveSheet($sheet_index = 0)
    {
        $this->phpExcel->setActiveSheetIndex($sheet_index);
    }
    
    /**
     * Sets the cell's value<br>
     * 
     * @param string $cell (A1)<br>
     * @param string $value<br>
     * 
     * @return void
     */
    public function setCellValue($cell, $value)
    {
        $this->phpExcel->getActiveSheet()->setCellValue($cell, $value);
    }
    
    /**
     * Sets the cell's value explicitely as string value<br>
     * <i>Ensures Excel won't force the cell to a strange datatype changing its value</i>
     * 
     * @param string $cell (A1)<br>
     * @param string $value<br>
     * 
     * @return void
     */
    public function setCellValueForceString($cell, $value)
    {
        $this->phpExcel->getActiveSheet()->setCellValueExplicit($cell, $value, PHPExcel_Cell_DataType::TYPE_STRING);
    }
    
    /**
     * Saves the file as Excel2007 format<br>
     * <i>This is quite system relevant as it requires an uploads folder in the root</i>
     * 
     * @param string $file_name The intended filename<br>
     * 
     * @return string The $file_name sent to the parameter
     */
    public function saveFile($file_name, $path_prefix = '')
    {
        $objWriter = PHPExcel_IOFactory::createWriter($this->phpExcel, "Excel2007");
        if(file_exists("uploads") && is_dir("uploads")) {
            $objWriter->save("{$path_prefix}uploads/$file_name");
        } else {
            $objWriter->save("{$path_prefix}../uploads/$file_name");
        }
        
//        $url = $_SERVER['SERVER_NAME'];
//        if($_SERVER['SERVER_NAME'] == 'develop1.lmsystem.co.za'){
//            $url = $url."/aon";
//        }
//        $url .= "/uploads/$file_name";
//        $objWriter->save($url);
        
        return $file_name;
    }
    
    /**** Regularly used functions ***/
    
    /**
     * Sets and styles column headers<br>
     * 
     * <b>Method 1 (Set column name and width)</b><br>
     * <i>
     * $arrayOfColumnDetails = array();<br>
     * $arrayOfColumnDetails[] = array("DATE AUTHORISED",12); //A<br>
     * $arrayOfColumnDetails[] = array("AUTHORISED BY",25);   //B<br>
     * </i>
     * <b>Method 2 (Set column name and width)</b><br>
     * <i>
     * $arrayOfColumnDetails = array();<br>
     * $arrayOfColumnDetails[] = "DATE AUTHORISED"; //A<br>
     * $arrayOfColumnDetails[] = "AUTHORISED BY";   //B<br>
     * </i>
     * @param array $arrayOfColumnDetails Array of Headings<br>
     * @param int $row Row where this array should be applied<br>
     * @param string $start_cell = "A" Starting Column<br>
     * @param string $style = "thinBorderBold" Style to be applied<br>
     * 
     * @return void
     */
    public function setColumnHeadersFromArray($arrayOfColumnDetails, $row, $start_cell="A", $style = "thinBorderBold")
    {
        $start_cell = strtoupper($start_cell);
        foreach($arrayOfColumnDetails as $columnDetails)
        {
            if (sizeof($columnDetails)>0)
            {
                $this->phpExcel->getActiveSheet()->setCellValue("{$start_cell}{$row}", $columnDetails[0]);
                $this->phpExcel->getActiveSheet()->getColumnDimension($start_cell)->setWidth($columnDetails[1]);
            } else
            {
                $this->phpExcel->getActiveSheet()->setCellValue("{$start_cell}{$row}", $columnDetails);
            }
            $this->phpExcel->getActiveSheet()->getStyle("{$start_cell}{$row}")->applyFromArray($this->getStyleFromDescription($style));
            $start_cell++;
        }
    }
    
    
    /********************************************** Standard formatting ***************************************/
    
    /**
     * Merges the range of cells<br>
     * 
     * @param string $range Range to be merged (E.g. "A1:A6" OR "A1:D1" OR "A1:D6")<br>
     * 
     * @return void
     */
    public function mergeCells($range)
    {
        $this->phpExcel->getActiveSheet()->mergeCells("$range");
    }
    
    /**
     * Set style of cell or range of cells<br>
     * 
     * @param string $range Range to be merged (E.g. "A1:A6" OR "A1:D1" OR "A1:D6")<br>
     * @param string $description Style description<br>
     * @param string $outline = 0 If 0 just format outline, else format each cell in range<br>
     * 
     * @return void
     */
    public function setStyle($range,$description,$outline = "0")
    {
        $range = strtoupper($range);
        if (strpos($range, ':') !== FALSE)//Called to format a range eg)A1:D9
        {
            if ($outline == "0")
            {
                list($cell_from, $cell_to) = explode(':',$range);
                list($col_from, $row_from) = preg_split('#(?<=[a-z])(?=\d)#i', $cell_from);
                list($col_to, $row_to) = preg_split('#(?<=[a-z])(?=\d)#i', $cell_to);

                for ($row = $row_from; $row <= $row_to; $row++)
                {
                    for ($col = $col_from; $col <= 'Z'; $col++) //Ensures to continue to AA, AB, AC ...
                    {
                        $this->phpExcel->getActiveSheet()->getStyle($col.$row)->applyFromArray($this->getStyleFromDescription($description));
                        if ($col==$col_to)
                            break;
                    }
                }
            } else//Just format outline
            {
                $this->phpExcel->getActiveSheet()->getStyle($range)->applyFromArray($this->getStyleFromDescription($description));
            }
        } else
        {   //Called to format a sibgle cell eg)A1
            $this->phpExcel->getActiveSheet()->getStyle($range)->applyFromArray($this->getStyleFromDescription($description));            
        }
    }
    
    /**
     * Set alignment of cell or range of cells<br>
     * 
     * @param string $range Range to be aligned (E.g. "A1:A6" OR "A1:D1" OR "A1:D6")<br>
     * @param string $description               "HL" - HORIZONTAL_LEFT 
     *                                     <br> "HC" - HORIZONTAL_CENTER 
     *                                     <br> "HR" - HORIZONTAL_RIGHT 
     *                                     <br> "VT" - VERTICAL_TOP 
     *                                     <br> "VC" - VERTICAL_CENTER 
     *                                     <br> "VB" - VERTICAL_BOTTOM 
     *                                     <br><i>Can be concatenated (E.g. "HL, VT")</i>
     * @return void
     */
    public function setAlignment($range, $description)
    {
        $description = strtolower($description);
        if (strpos($description, 'hr') !== FALSE)//HORIZONTAL_RIGHT
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
        }
        if (strpos($description, 'hl') !== FALSE)//HORIZONTAL_LEFT
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        }
        if (strpos($description, 'hc') !== FALSE)//HORIZONTAL_CENTER
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        }
        if (strpos($description, 'hc') !== FALSE)//HORIZONTAL_CENTER
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        }
        if (strpos($description, 'vb') !== FALSE)//VERTICAL_RIGHT
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_BOTTOM);
        }
        if (strpos($description, 'vt') !== FALSE)//VERTICAL_LEFT
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);
        }
        if (strpos($description, 'vc') !== FALSE)//VERTICAL_CENTER
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
        }
        if (strpos($description, 'wt') !== FALSE)//WrapText
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getAlignment()->setWrapText(true);
        }
    }
    
    /**
     * Sets the font size of cell or range of cells<br>
     * 
     * @param string $range Range (E.g. "A1:A6" OR "A1:D1" OR "A1:D6")<br>
     * @param int $size Desired font size<br>
     * 
     * @return void
     */
    public function setFontSize($range, $size)
    {
        $this->phpExcel->getActiveSheet()->getStyle($range)->getFont()->setSize($size);
    }
    
    /**
     * Sets the height of specified row<br>
     * 
     * @param int $row Desired row<br>
     * @param int $height Desired row height<br>
     * 
     * @return void
     */
    public function setRowHeight($row, $height)
    {
        $this->phpExcel->getActiveSheet()->getRowDimension($row)->setRowHeight($height);
    }
    
    /**
     * Sets the height of range of rows<br>
     * 
     * @param int $row_from Starting row<br>
     * @param int $row_to Ending row<br>
     * @param int $height Desired row height<br>
     * 
     * @return void
     */
    public function setRowHeightByRange($row_from, $row_to, $height)
    {
        for($i=$row_from; $i<$row_to; $i++) {
            $this->phpExcel->getActiveSheet()->getRowDimension($i)->setRowHeight($height);
        }
    }
    
    //Font Styles
    /**
     * Sets the font style of cell or range of cells<br>
     * 
     * @param string $range Range (E.g. "A1:A6" OR "A1:D1" OR "A1:D6")<br>
     * @param string $description "B" - Bold
     *                       <br> "I" - Italic
     *                       <br> "U" - Underline
     *                       <br><i>Can be concatenated (E.g. "B, I, U")</i>
     * 
     * @return void
     */
    public function setFontStyle($range,$description)
    {
        $description = strtolower($description);
        if (strpos($description, 'b') !== FALSE)//bold
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getFont()->setBold(true);
        }
        if (strpos($description, 'i') !== FALSE)//italic
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getFont()->setItalic(true);
        }
        if (strpos($description, 'u') !== FALSE)//underline
        {
            $this->phpExcel->getActiveSheet()->getStyle($range)->getFont()->setUnderline(true);
        }
    }
    
    /**
     * Set column widths by array<br>
     * 
     * @param array $arrayOfColumnSizes Array of int values<br>
     * @param string $start_cell Column to start<br>
     * 
     * @return void
     */
    public function setColumnWidthsFromArray($arrayOfColumnSizes, $start_cell="A")
    {
        $start_cell = strtoupper($start_cell);
        foreach($arrayOfColumnSizes as $columnSize)
        {
            $this->phpExcel->getActiveSheet()->getColumnDimension($start_cell)->setWidth($columnSize);
            $start_cell++;
        }
    }
    
    /**
     * Set Number format for range (Default is Currency)<br>
     * <i>Any Excel format formula can be used here</i>
     * 
     * @param string $range Range (E.g. "A1:A6" OR "A1:D1" OR "A1:D6")<br>
     * @param string $description "yyyy-mm-dd hh:mm" - Date with time
     *                       <br> "0.00%" - Percentage
     *                       <br> "R # ##0;[Red]R -# ##0" - Currency wityh negatives in red
     *                       <br> "_ R * # ##0.00_ ;_ R * -# ##0.00_ ;_ R * "-"??_ ;_ @_ " - Accounting
     * 
     * @return void
     */
    function setNumberFormat($range, $description = "currency")
    {
        switch ($description)
        {
            case "currency":
                $pattern = '"R "#,##0.00_-';
                break;
            default:
                $pattern = $description;
                break;
        }
        $this->phpExcel->getActiveSheet()->getStyle($range)->getNumberFormat()->setFormatCode($pattern);
    }
    
    private function getStyleFromDescription($description)
    {
        $description = strtolower($description);
        switch($description){
            case "fileheading": 
                return $this->fileHeading;
                break;
            case "thinborderbold":
                return $this->thinBorderBold;
                break;
            case "stylemediumblackborderoutline":
                return $this->styleMediumBlackBorderOutline;
                break;
            case "stylethinblackborderoutline":
                return $this->styleThinBlackBorderOutline;
                break;
            case "styletotal":
                return $this->styleTotal;
                break;
            case "styletotallabel":
                return $this->styleTotalLabel;
                break;
            case "boldtext":
                return $this->boldText;
                break;
            case "normalcell":
                return $this->normalCell;
                break;
            default:
                return $this->normalCell;
                break;
        }
    }
            
    /************************************************ Styles ************************************************/
    
    /**
     * <b>File Heading</b><br>
     * <br><i>Font -> BOLD
     * <br><i>Alignment -> HORIZONTAL_CENTER
     * <br><i>Borders -> Top -> BORDER_THIN
     * <br><i>Fill -> Type -> FILL_GRADIENT_LINEAR
     * <br><i>Fill -> startcolor -> FFA0A0A0
     * <br><i>Fill -> endcolor -> FFFFFFFF
     * <br><i>Fill -> Rotation -> 90
     */
    public $fileHeading = array(
			"font"    => array(
				"bold"      => true
			),
			"alignment" => array(
				"horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
			),
			"borders" => array(
				"top"     => array(
 					"style" => PHPExcel_Style_Border::BORDER_THIN
 				)
			),
			"fill" => array(
	 			"type"       => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
	  			"rotation"   => 90,
	 			"startcolor" => array(
	 				"argb" => "FFA0A0A0"
	 			),
	 			"endcolor"   => array(
	 				"argb" => "FFFFFFFF"
	 			)
	 		)
		);
    
    /**
     * <b>Thin Border Bold</b><br>
     * <br><i>Font -> BOLD
     * <br><i>Alignment -> HORIZONTAL_CENTER
     * <br><i>Borders -> Outline -> BORDER_MEDIUM
     */
    public $thinBorderBold = array(
                            "font"    => array(
                                    "bold"      => true
                            ),
                            "alignment" => array(
                                    "horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                            ),
                            "borders" => array(
                                    "outline"     => array(
                                            "style" => PHPExcel_Style_Border::BORDER_MEDIUM
                                    )
                            )
                    );

    /**
     * <b>Medium Black Border Outline</b><br>
     * <br><i>Borders -> Outline -> BORDER_MEDIUM
     * <br><i>Borders -> Outline -> FF000000
     */
    public $styleMediumBlackBorderOutline = array(
            'borders' => array(
                    'outline' => array(
                            'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                            'color' => array('argb' => 'FF000000'),
                    ),
            ),
    );
    
    /**
     * <b>Thin Black Border Outline</b><br>
     * <br><i>Borders -> Outline -> BORDER_THIN
     * <br><i>Borders -> Outline -> FF000000
     */
    public $styleThinBlackBorderOutline = array(
            'borders' => array(
                    'outline' => array(
                            'style' => PHPExcel_Style_Border::BORDER_THIN,
                            'color' => array('argb' => 'FF000000'),
                    ),
            ),
    );

    /**
     * <b>Total</b><br>
     * <br><i>Borders -> Bottom -> BORDER_DOUBLE
     */
    public $styleTotal = array(
              'borders' => array(
                        "bottom" => array(
                                  'style' => PHPExcel_Style_Border::BORDER_DOUBLE
                        )
              )
    );

    /**
     * <b>Total Label</b><br>
     * <br><i>Font -> BOLD
     * <br><i>Alignment -> HORIZONTAL_RIGHT
     */
    public $styleTotalLabel = array(
              "font"    => array(
                        "bold" => true
              ),
              'alignment' => array(
                        "horizontal" => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT
              )
    );

    public $boldText = array("font"=>array("bold"=>true));
    public $normalCell = array("alignment"=>array("horizontal"=>PHPExcel_Style_Alignment::HORIZONTAL_LEFT));
}
}
?>