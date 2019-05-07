<?php

namespace app\modules\news\controllers;

use Yii;
use app\models\Article;
use app\models\ArticleSearh;
use yii\web\Controller;
use yii\web\NotFoundHttpException;
use yii\filters\VerbFilter;
use yii\data\Pagination;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

/**
 * DefaultController implements the CRUD actions for Article model.
 */
class DefaultController extends Controller
{
    /**
     * {@inheritdoc}
     */
    public function behaviors()
    {
        return [
            'verbs' => [
                'class' => VerbFilter::className(),
                'actions' => [
                    'delete' => ['POST'],
                ],
            ],
        ];
    }

    /**
     * Lists all Article models.
     * @return mixed
     */
    public function actionIndex()
    {
        $searchModel = new ArticleSearh();
        $dataProvider = $searchModel->search(Yii::$app->request->queryParams);
        //$query = Article::find();
        //$count = $query->count();
        //$pagination = new Pagination(['totalCount' => $count]);
        //var_dump($pagination);exit;111

        return $this->render('index', [
            'searchModel' => $searchModel,
            'dataProvider' => $dataProvider,
        ]);
    }

    /**
     * Displays a single Article model.
     * @param integer $id
     * @return mixed
     * @throws NotFoundHttpException if the model cannot be found
     */
    public function actionView($id)
    {
        return $this->render('view', [
            'model' => $this->findModel($id),
        ]);
    }

    /**
     * Creates a new Article model.
     * If creation is successful, the browser will be redirected to the 'view' page.
     * @return mixed
     */
    public function actionCreate()
    {
        $model = new Article();

        if ($model->load(Yii::$app->request->post()) && $model->save()) {
            return $this->redirect(['view', 'id' => $model->id]);
        }

        return $this->render('create', [
            'model' => $model,
        ]);
    }

    /**
     * Updates an existing Article model.
     * If update is successful, the browser will be redirected to the 'view' page.
     * @param integer $id
     * @return mixed
     * @throws NotFoundHttpException if the model cannot be found
     */
    public function actionUpdate($id)
    {
        $model = $this->findModel($id);

        if ($model->load(Yii::$app->request->post()) && $model->save()) {
            return $this->redirect(['view', 'id' => $model->id]);
        }

        return $this->render('update', [
            'model' => $model,
        ]);
    }

    /**
     * Deletes an existing Article model.
     * If deletion is successful, the browser will be redirected to the 'index' page.
     * @param integer $id
     * @return mixed
     * @throws NotFoundHttpException if the model cannot be found
     */
    public function actionDelete($id)
    {
        $this->findModel($id)->delete();

        return $this->redirect(['index']);
    }

    /**
     * Finds the Article model based on its primary key value.
     * If the model is not found, a 404 HTTP exception will be thrown.
     * @param integer $id
     * @return Article the loaded model
     * @throws NotFoundHttpException if the model cannot be found
     */
    protected function findModel($id)
    {
        if (($model = Article::findOne($id)) !== null) {
            return $model;
        }

        throw new NotFoundHttpException('The requested page does not exist.');
    }

    public function actionGetFileInfo($name)
    {
        $path = __DIR__."\office\\$name";
        if (file_exists($path)) {
            $handle = fopen($path, "r");
            $size = filesize($path);
            $contents = fread($handle, $size);
            $SHA256 = base64_encode(hash('sha256', $contents, true));
            $json = array(
                'BaseFileName' => $name,
                'OwnerId' => 'admin',
                'Size' => $size,
                'SHA256' => $SHA256,
                'Version' => '222888822',
                "AllowExternalMarketplace"=>true,
                "UserCanWrite"=>true,
                "SupportsUpdate"=>true,
                "SupportsLocks"=>true
            );
            header('Content-Type: application/json');
            echo json_encode($json);
        } else {
            echo json_encode(array());
        }
    }

    public function actionGetFile($name) {
        if(Yii::$app->request->isPost){
            $this->actionPutfile($name);
        }else{
            $path = __DIR__."\office\\$name";
            if (file_exists($path)) {
                header('Content-Description: File Transfer');
                header('Content-Type: application/octet-stream');
                header('Content-Disposition: attachment; filename=' . basename($path));
                header('Expires: 0');
                header('Cache-Control: must-revalidate');
                header('Pragma: public');
                header('Content-Length: ' . filesize($path));
                readfile($path);
                exit;
            }
        }
        
    }

    public function actionPutfile($name)
    {
        $path = __DIR__."\office\\$name";
        $contents = file_get_contents('php://input');
        if (file_exists($path)) {
            file_put_contents($path, $contents);
        }
        echo $contents;
    }

    public function actionTest123()
    {
        echo urlencode('http://192.168.110.1:8002/wopihost/wopi/files/test.docx');
    }

    public function init(){
        $this->enableCsrfValidation = false;
    }

    public function actionHtmlToDocx()
    {
        $htmlPath = "D:\\testcsharp\\testfile\\123.html";
        $time = time();
        $docPath = "D:\\testcsharp\\testfile\\123-{$time}.docx";
        $cmd = <<<EOF
\$htmlPath = '$htmlPath';
\$docPath = '$docPath';
\$wordApp = New-Object -ComObject Word.Application;
\$document = \$wordApp.Documents.Open(\$htmlPath,\$false);
\$document.SaveAs([ref] \$docPath, [ref] 16);
\$document.Close();
\$wordApp.Quit();
EOF;
        $cmd = implode("", explode("\n", $cmd));
        $f=shell_exec('powershell ' . $cmd);
        var_dump($f);
    }

    public function actionDocxToHtml()
    {
        $docPath = "D:\\testcsharp\\testfile\\123.docx";
        $time = time();
        $htmlPath = "D:\\testcsharp\\testfile\\123-{$time}.html";
        $cmd = <<<EOF
            \$docPath = '$docPath';
            \$htmlPath = '$htmlPath';
            \$wordApp = New-Object -ComObject Word.Application;
            \$document = \$wordApp.Documents.Open(\$docPath);
            \$document.SaveAs([ref] \$htmlPath, [ref] 10);
            \$document.Close();
            \$wordApp.Quit();
EOF;
        $cmd = implode("", explode("\n", $cmd));
        $f=shell_exec('powershell ' . $cmd);
        var_dump($f);
    }

    public function actionTest()
    {
        $path = "D:\\test\\test.xlsx";

        $reader = IOFactory::createReader('Xlsx');
        $spreadsheet = $reader->load($path);

        //$helper->log('Add new data to the template');
        $data = '{"List":{"2016":{"Jan":0,"Feb":0,"Mar":0,"Apr":0,"May":0,"Jun":0,"Jul":6841.26,"Aug":21207.91,"Sep":21207.91,"Oct":21207.91,"Nov":21207.91,"Dec":21207.91},"2017":{"Jan":55414.21,"Feb":127247.44,"Mar":127247.44,"Apr":127247.44,"May":127247.44,"Jun":127247.44,"Jul":129352.44,"Aug":133772.95,"Sep":133772.95,"Oct":133772.95,"Nov":133772.95,"Dec":133772.95},"2018":{"Jan":133772.95,"Feb":133772.95,"Mar":133772.95,"Apr":133772.95,"May":133772.95,"Jun":133772.95,"Jul":135877.95,"Aug":140298.46,"Sep":140298.46,"Oct":140298.46,"Nov":140298.46,"Dec":140298.46},"2019":{"Jan":140298.46,"Feb":140298.46,"Mar":140298.46,"Apr":126493.62,"May":139034.51,"Jun":139034.51,"Jul":141120.55,"Aug":145501.23,"Sep":145501.23,"Oct":145501.23,"Nov":145501.23,"Dec":145501.23},"2020":{"Jan":145501.23,"Feb":145501.23,"Mar":145501.23,"Apr":145501.23,"May":145501.23,"Jun":145501.23,"Jul":147587.27,"Aug":151967.95,"Sep":151967.95,"Oct":151967.95,"Nov":151967.95,"Dec":151967.95},"2021":{"Jan":151967.95,"Feb":151967.95,"Mar":151967.95,"Apr":151967.95,"May":151967.95,"Jun":151967.95,"Jul":102946.04,"Aug":0,"Sep":0,"Oct":0,"Nov":0,"Dec":0}},"Total":7732030.22,"Count":{"6841.26":{"months":0.32,"totalRental":6841.26},"21207.91":{"months":5,"totalRental":106039.55},"55414.21":{"months":1,"totalRental":55414.21},"127247.44":{"months":5,"totalRental":636237.2},"129352.44":{"months":1,"totalRental":129352.44},"133772.95":{"months":11,"totalRental":1471502.45},"135877.95":{"months":1,"totalRental":135877.95},"140298.46":{"months":8,"totalRental":1122387.68},"126493.62":{"months":1,"totalRental":126493.62},"139034.51":{"months":2,"totalRental":278069.02},"141120.55":{"months":1,"totalRental":141120.55},"145501.23":{"months":11,"totalRental":1600513.53},"147587.27":{"months":1,"totalRental":147587.27},"151967.95":{"months":11,"totalRental":1671647.45},"102946.04":{"months":0.68,"totalRental":102946.04}},"Term":60,"CountTotalRental":7732030.22,"CountEffRental":231.15,"EffRental":231.15,"taxChangeTable":{"2018-Jun":{"before":133772.95,"after":132567.79},"2018-Jul":{"before":135877.95,"after":134653.82},"2018-Aug":{"before":140298.46,"after":139034.51},"2018-Sep":{"before":140298.46,"after":139034.51},"2018-Oct":{"before":140298.46,"after":139034.51},"2018-Nov":{"before":140298.46,"after":139034.51},"2018-Dec":{"before":140298.46,"after":139034.51},"2019-Jan":{"before":140298.46,"after":139034.51},"2019-Feb":{"before":140298.46,"after":139034.51},"2019-Mar":{"before":140298.46,"after":139034.51},"total":{"before":1392038.58,"after":1379497.69},"2019-Apr":{"before":140298.46,"after":126493.62}}}';
        $rent = [228.246525,239.951475,251.656425,263.361375,275.066325];
        $lease_area = 557.5;
        $freePeriod = 5;
        $lcd = '2016.07.22';
        $data = json_decode($data,true);

        //每年单价
        $spreadsheet->getActiveSheet()->setCellValue('B2', $rent[0]);
        $spreadsheet->getActiveSheet()->setCellValue('B3', $rent[1]);
        $spreadsheet->getActiveSheet()->setCellValue('B4', $rent[2]);
        $spreadsheet->getActiveSheet()->setCellValue('B5', $rent[3]);
        $spreadsheet->getActiveSheet()->setCellValue('B6', $rent[4]);

        //免租期
        $spreadsheet->getActiveSheet()->setCellValue('D2', $freePeriod);

        //面积
        $spreadsheet->getActiveSheet()->setCellValue('F2', $lease_area);

        //起租日期
        $spreadsheet->getActiveSheet()->setCellValue('F3', $lcd);
        
        //total
        $spreadsheet->getActiveSheet()->setCellValue('B10', $data['Total']);
        //eff.rental
        $spreadsheet->getActiveSheet()->setCellValue('D10', $data['EffRental']);

        //term
        $spreadsheet->getActiveSheet()->setCellValue('B16', $data['Term']);
        //CountTotalRental
        $spreadsheet->getActiveSheet()->setCellValue('C16', $data['CountTotalRental']);
        //CountEffRental
        $spreadsheet->getActiveSheet()->setCellValue('C17', $data['CountEffRental']);

        //税率调整相关
        //$spreadsheet->getActiveSheet()->setCellValue('C20', $data['taxChangeTable']['total']['before']);
        //$spreadsheet->getActiveSheet()->setCellValue('C21', $data['taxChangeTable']['total']['after']);
        //$spreadsheet->getActiveSheet()->setCellValue('C22', $data['taxChangeTable']['total']['before'] - $data['taxChangeTable']['total']['after']);

        //输出试算结果
        $recordBaseRow = 9;
        $i = 0;
        foreach ($data['List'] as $year => $months) {
            $row = $recordBaseRow + $i;
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);

            $spreadsheet->getActiveSheet()->setCellValue('A' . $row, $year)
                ->setCellValue('B' . $row, $months['Jan'])
                ->setCellValue('C' . $row, $months['Feb'])
                ->setCellValue('D' . $row, $months['Mar'])
                ->setCellValue('E' . $row, $months['Apr'])
                ->setCellValue('F' . $row, $months['May'])
                ->setCellValue('G' . $row, $months['Jun'])
                ->setCellValue('H' . $row, $months['Jul'])
                ->setCellValue('I' . $row, $months['Aug'])
                ->setCellValue('J' . $row, $months['Sep'])
                ->setCellValue('K' . $row, $months['Oct'])
                ->setCellValue('L' . $row, $months['Nov'])
                ->setCellValue('M' . $row, $months['Dec']);
            $i++;
        }
        $spreadsheet->getActiveSheet()->removeRow($row + 1, 1);
        unset($row);

        //输出统计结果
        $countBaseRow = 15 + $i - 1;
        $j = 0;
        foreach ($data['Count'] as $price => $value) {
            $row = $countBaseRow + $j;
            $spreadsheet->getActiveSheet()->insertNewRowBefore($row, 1);
            $spreadsheet->getActiveSheet()->setCellValue('A' . $row, $price)
                ->setCellValue('B' . $row, $value['months'])
                ->setCellValue('C' . $row, $value['totalRental']);
                $j++;
        }
        $spreadsheet->getActiveSheet()->removeRow($row + 1, 1);
        unset($row);

        //税率表填充
        $taxTableBaseRow = 19 + $i - 1 + $j - 1;
        $k = 66;//ASCII码大写字母B是65 ord('B')可以得出
        $tableStart = chr($k) . $taxTableBaseRow;
        foreach ($data['taxChangeTable'] as $date => $value) {

            $spreadsheet->getActiveSheet()->setCellValue(chr($k) . $taxTableBaseRow, $date == 'total' ? '合计' : $date)
                ->setCellValue(chr($k) . ($taxTableBaseRow + 1), $value['before'])
                ->setCellValue(chr($k) . ($taxTableBaseRow + 2), $value['after'])
                ->setCellValue(chr($k) . ($taxTableBaseRow + 3), $value['before'] - $value['after']);
            $k++;
        }
        $tableEnd = chr($k - 1) . ($taxTableBaseRow + 3);

        //格式化好税率表
        $spreadsheet->getActiveSheet()->getStyle($tableStart.':'.$tableEnd)->applyFromArray(
            [
                'borders' =>
                [
                    'bottom' => ['borderStyle' => Border::BORDER_THIN],
                    'right' => ['borderStyle' => Border::BORDER_THIN],
                    'left' => ['borderStyle' => Border::BORDER_THIN],
                    'top' => ['borderStyle' => Border::BORDER_THIN],
                    'vertical' => ['borderStyle' => Border::BORDER_THIN],
                    'horizontal' => ['borderStyle' => Border::BORDER_THIN]
                ],
                'font' =>
                [
                    'name'=> 'Calibri',
                    'size'=> 12
                ],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical' => Alignment::VERTICAL_CENTER,
                    'wrapText' => true,
                ],
            ]
        );
        $spreadsheet->setActiveSheetIndex(0);

        // Write documents
        //$filepath = $this->getFilename(__FILE__, mb_strtolower('xlsx'));
        $time = time();
        $filepath = "D:\\test\\test-{$time}.xlsx";
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($filepath);
    }
}