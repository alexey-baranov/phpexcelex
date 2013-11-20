<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

namespace PHPExcelEx\Reader;

/**
 * Читает объект PDOStatement в Excel 
 * 
 * $locale = 'pt_br';
 * $validLocale = PHPExcel_Settings::setLocale($locale);
 * if (!$validLocale) {
 *	echo 'Unable to set locale to '.$locale." - reverting to en_us<br />\n";
 * }
 *
 * @author Администратор
 */
class PDO implements \PHPExcel_Reader_IReader{
    function canRead($source) {
        if ($source instanceof \PDOStatement || $source instanceof \Doctrine\DBAL\Statement){
            return true;
        }
    }
    /**
     *
     * @param PDOStatement $source
     * @param PHPExcel $excel есил нужно добавить sheet к PHPExcel
     * @return PHPExcel
     */
    public function load($source, $excel= null) {
        /* @var $source \Doctrine\DBAL\Statement */
        if ($excel) {
            $sheet= $excel->createSheet();
        }
        else {
            $excel= new \PHPExcel();
            $sheet= $excel->getActiveSheet();
        }
        
        //filling headers
        for($EACH_COLUMN=0; $EACH_COLUMN<$source->columnCount(); $EACH_COLUMN++) {
            $eachColumnMeta= $source->getColumnMeta($EACH_COLUMN);
            $sheet->setCellValueByColumnAndRow($EACH_COLUMN, 1, $eachColumnMeta['name']);
        }
        
        //filling data
        $EACH_ROW=2;
        while($eachRowAsArray=$source->fetch(\PDO::FETCH_NUM)) {
            for($EACH_COLUMN=0; $EACH_COLUMN<$source->columnCount(); $EACH_COLUMN++) {
                $eachColumnMeta= $source->getColumnMeta($EACH_COLUMN);
                switch($eachColumnMeta['native_type']){
                    case 'timestampt':
                    case 'timestamptz':
                        $eachColumnValue= new \DateTime($eachRowAsArray[$EACH_COLUMN]);
                        $sheet->setCellValueByColumnAndRow($EACH_COLUMN, $EACH_ROW, $eachColumnValue->format("d.m.Y H:i"));
                        break;
                    default:
                        $sheet->setCellValueByColumnAndRow($EACH_COLUMN, $EACH_ROW, $eachRowAsArray[$EACH_COLUMN]);
                        break;
                }
            }
            $EACH_ROW++; //s
        }
        return $excel;
    }
}