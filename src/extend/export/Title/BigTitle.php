<?php


namespace MillionMile\PHPSpreadsheet\extend\export\Title;

use MillionMile\PHPSpreadsheet\core\Common;
use MillionMile\PHPSpreadsheet\core\Export;
use MillionMile\PHPSpreadsheet\extend\export\Style\StyleConfig;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

class BigTitle
{
    private $export;
    private $mergeCount;
    private $styleObj;
    private $defaultStyle;
    private $rowHeight;

    public function __construct(int $mergeCount = 1)
    {
        $this->export = Export::getInstance();
        $this->mergeCount = $mergeCount;

        //设置默认样式
        $this->defaultStyle = [
            'font' => [
                'bold' => true,
                'size' => 16
            ],
            'horizontal' => Alignment::HORIZONTAL_CENTER, // 水平居中
            'vertical' => Alignment::VERTICAL_CENTER,     // 垂直居中
            'wrapText' => true

        ];
        $this->rowHeight = 48;
    }

    /**
     * @title 人工设置样式对象
     * @param StyleConfig $styleObj
     * @author millionmile
     * @time 2020/07/06 12:09
     */
    public function setStyleObj(StyleConfig $styleObj)
    {
        $this->styleObj = $styleObj;
    }


    /**
     * @title 获取样式对象，如果不存在，则自己创建
     * @author millionmile
     * @time 2020/07/06 12:10
     */
    private function getStyleObj()
    {

    }


    public function setBigTitle(string $titleStr,$style)
    {
        $this->getStyleObj();
        $sheet =& $this->export->getActiveSheet();
        //插入一行
        try {
            $this->export->insertRows(1, 1);
            //在首行加入大标题
            $sheet->setCellValue('A1', $titleStr);
            $endCol = Common::getColName($this->mergeCount);
            $sheet->mergeCells('A1:' . $endCol . '1');
            if ($style){
                $sheet->getStyle('A1:' . $endCol . '1')->applyFromArray($style['style']);
                $sheet->getRowDimension(1)->setRowHeight($style['row_height']);
            }else{
                $sheet->getStyle('A1:' . $endCol . '1')->applyFromArray($this->defaultStyle);
                $sheet->getRowDimension(1)->setRowHeight($this->rowHeight);
            }


        } catch (Exception $e) {
            return $e->getMessage();
        }
        return true;
    }
}
