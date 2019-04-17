<?php

namespace app\modules\news\controllers;

use Yii;
use app\models\Article;
use app\models\ArticleSearh;
use yii\web\Controller;
use yii\web\NotFoundHttpException;
use yii\filters\VerbFilter;
use yii\data\Pagination;

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
}
