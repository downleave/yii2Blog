<?php

namespace app\controllers;

use Yii;
use yii\filters\AccessControl;
use yii\web\Controller;
use yii\web\Response;
use yii\filters\VerbFilter;
use app\models\LoginForm;
use app\models\ContactForm;
use PhpOffice\PhpWord;

class SiteController extends Controller
{
    /**
     * {@inheritdoc}
     */
    public function behaviors()
    {
        return [
            'access' => [
                'class' => AccessControl::className(),
                'only' => ['logout'],
                'rules' => [
                    [
                        'actions' => ['logout'],
                        'allow' => true,
                        'roles' => ['@'],
                    ],
                ],
            ],
            'verbs' => [
                'class' => VerbFilter::className(),
                'actions' => [
                    'logout' => ['post'],
                ],
            ],
        ];
    }

    /**
     * {@inheritdoc}
     */
    public function actions()
    {
        return [
            'error' => [
                'class' => 'yii\web\ErrorAction',
            ],
            'captcha' => [
                'class' => 'yii\captcha\CaptchaAction',
                'fixedVerifyCode' => YII_ENV_TEST ? 'testme' : null,
            ],
        ];
    }

    /**
     * Displays homepage.
     *
     * @return string
     */
    public function actionIndex()
    {
        return $this->redirect(array('/news/default'));
        //return $this->render('index');
    }

    /**
     * Login action.
     *
     * @return Response|string
     */
    public function actionLogin()
    {
        if (!Yii::$app->user->isGuest) {
            return $this->goHome();
        }

        $model = new LoginForm();
        if ($model->load(Yii::$app->request->post()) && $model->login()) {
            return $this->goBack();
        }

        $model->password = '';
        return $this->render('login', [
            'model' => $model,
        ]);
    }

    /**
     * Logout action.
     *
     * @return Response
     */
    public function actionLogout()
    {
        Yii::$app->user->logout();

        return $this->goHome();
    }

    /**
     * Displays contact page.
     *
     * @return Response|string
     */
    public function actionContact()
    {
        $model = new ContactForm();
        if ($model->load(Yii::$app->request->post()) && $model->contact(Yii::$app->params['adminEmail'])) {
            Yii::$app->session->setFlash('contactFormSubmitted');

            return $this->refresh();
        }
        return $this->render('contact', [
            'model' => $model,
        ]);
    }

    /**
     * Displays about page.
     *
     * @return string
     */
    public function actionAbout()
    {
        return $this->render('about');
    }

    public function actionTest()
    {
        // Read contents
        // echo date('H:i:s'), " Reading contents from `{$source}`", EOL;
        // $phpWord = \PhpOffice\PhpWord\IOFactory::load($source,$type);
        // var_dump($phpWord);die;
        // $phpWord = new \PhpOffice\PhpWord\Reader\Word2007();
        // $template = $phpWord->load($source);
        //var_dump($template->getSections()[0]->getelements()[0]->getText());

        // foreach ($template->getSections() as $sections_key => $sections_value) {
        //     foreach ($sections_value->getelements() as $elements_key => $elements_value) {
        //         var_dump($elements_value->getText());
        //     }
        // }
        //var_dump($template->getSections());
        //var_dump($phpWord->getSections()[0]->getelements()[0]->getText());


        // $template = new \PhpOffice\PhpWord\TemplateProcessor($source);
        // var_dump($template);die;
        // $tmp_file = 'test_tmp_file.txt';
        // file_put_contents($tmp_file, $template);
        // $a=file_get_contents($tmp_file);
        // var_dump($a);die;
        // $xmldata = $template->processSegment('CASE 2', 'w:p', \PhpOffice\PhpWord\TemplateProcessor::SEARCH_AROUND, 1, 'MainPart', function(&$xmlSegment, &$segmentStart, &$segmentEnd, &$part){
        //     $segmentStart = strpos($part, '<w:sdtPr', $segmentEnd + 1);
        //     if (!$segmentStart) dd("FATAL: nothing found!");
        //     $segmentEnd = strpos($part, '</w:sdtPr>', $segmentStart + 8);
        //     $xmlSegment = substr($part, $segmentStart, ($segmentEnd - $segmentStart));
        //     return false; # only getSegment
        // });

        // $p = xml_parser_create();
        // xml_parse_into_struct($p, $xmldata, $vals, $index);
        // xml_parser_free($p);
        // $date = $vals[2]{'attributes'}{'W:FULLDATE'}; #  "date" = "2017-11-11T00:00:00Z"
        // $type = $vals[5]{'attributes'}{'W:VAL'}; #  "type" = "dateTime"

        // $hash = [];
        // array_walk($vals, function(&$item) use(&$hash){
        //     if(array_key_exists('attributes',$item) && array_key_exists('tag',$item)){
        //         $hash[ $item['tag'] ] = array_values($item['attributes'])[0] ;
        //     }
        // });
        //$template->setValue(['购物艺术中心'],['替换测试test']);
        // var_dump(compact('vals','index','date', 'type', 'hash'));

        // preg_match_all('/(购物艺术中心)/', $template->tempDocumentMainPart, $matchsB);//勋总产品原型相对路径"/"开头的都匹配
        
        // if ($matchsB[1]) {
        //     $tmpReplaceValue=[];
        //     foreach ($matchsB[1] as $field) {
        //         $tmpReplaceValue[]="替换测试test";
        //     }

        //     $template->tempDocumentMainPart=str_replace($matchsB[1], $tmpReplaceValue, $template->tempDocumentMainPart);

        // }

        // $export_file='test20181123__01.docx';
        // $template->saveAs($export_file);
        //var_dump($matchsB);

        // $sections = $phpWord->getSections();

        // foreach ($sections as $section) {
        //   foreach ($section->getElements() as $element) {
        //      $string = $element->gettext();
        //      var_dump($string);
        //      exit;
        //   }
        // }

        // $phpWord = new \PhpOffice\PhpWord\Reader\Word2007();
        // $result=$phpWord->load($source);
        // var_dump($result);

        // $xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord,'Word2007');
        // //声明临时html文件
        // $tmpFile = 'test20181123__01.docx';

        // // //将$xmlWriter的值写入临时html文件
        // $xmlWriter->save($tmpFile);

        // //获取临时临时html文件中的内容
        // $content = file_get_contents($tmpFile);

        // //删除临时html文件
        // unlink($tmpFile);

        // //输出读取的内容
        // var_dump($content);


        // $name = basename(__FILE__, '.php');
        // $source = "resources/{$name}.doc";
        // echo date('H:i:s'), " Reading contents from `{$source}`", EOL;
        // $phpWord = \PhpOffice\PhpWord\IOFactory::load($source, 'MsDoc');
        // // (Re)write contents
        // $writers = array('Word2007' => 'docx', 'ODText' => 'odt', 'RTF' => 'rtf');
        // foreach ($writers as $writer => $extension) {
        //     echo date('H:i:s'), " Write to {$writer} format", EOL;
        //     $xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, $writer);
        //     $xmlWriter->save("{$name}.{$extension}");
        //     rename("{$name}.{$extension}", "results/{$name}.{$extension}");
        // }
        /*-------------------------测试html生成docx文档---------------------------------------*/

        /*        echo date('H:i:s') , ' Create new PhpWord object',"\n";
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $phpWord->addParagraphStyle('Heading2', array('alignment' => 'center'));
        $section = $phpWord->addSection();

        $html = self::getTestHtml();
        //var_dump($html);die;

        // $doc = new \DOMDocument();
        // $doc->loadHTML($html);
        // $doc->saveHTML();
        // \PhpOffice\PhpWord\Shared\Html::addHtml($section, $doc->saveHtml(),true);

        \PhpOffice\PhpWord\Shared\Html::addHtml($section, $html);     
        $filename='test_html_to_word2.docx';
        //\common\extensions\office\OfficeCommon::instance()->write($phpWord, $filename, ['Word2007' => 'docx']);
        $phpWord->save($filename, 'Word2007');*/
        /*-------------------------测试html生成docx文档---------------------------------------*/

        //$html=self::getTestHtml();
        //file_put_contents('test_html_to_word2.html', $html);

        // $doc = new \DOMDocument();
        // $doc->load("xml.xml");
        // var_dump($doc);
    }

    public function actionTestlala()
    {
        $path = 'D:\\test\\';
        $source = $path.'3_1.docx';
        $phpWord = \PhpOffice\PhpWord\IOFactory::load($source,'Word2007');
        $xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, "HTML");
        //声明临时html文件
        $tmpFile = $path.'3_1-'.time().'.html';

        // //将$xmlWriter的值写入临时html文件
        $ret = $xmlWriter->save($tmpFile);

        //$htmlOBJ = \PhpOffice\PhpWord\IOFactory::load($tmpFile,'HTML');
        //$docxWriter = \PhpOffice\PhpWord\IOFactory::createWriter($htmlOBJ,'Word2007');
        //$docFile = $path.'4-2.docx';
        //$ret = $htmlWriter->save($docFile);

        echo 'success';
    }

    public function actionTestlala2()
    {
        $path = 'D:\\test\\';
        $source = $path.'4.docx';
        $PHPWord = new PHPWord\TemplateProcessor($source);
        //$tempPlete = $PHPWord->loadTemplate($source);
        $PHPWord->setValue('jiafang','广州新御房地产');
        $PHPWord->setValue('year','2019');
        $PHPWord->setValue('month','12');
        $PHPWord->setValue('day','27');
        $PHPWord->setValue('price','1000');
        $PHPWord->setValue('price2','10000');
        $PHPWord->setValue('pricechina','壹万整');
        $tmpFile = $path.'4-2-1-'.time().'.docx';

        $PHPWord->saveAs($tmpFile);
        echo 'success';
    }

    public function actionTestlala3()
    {
        //$writers = array('Word2007' => 'docx', 'ODText' => 'odt', 'RTF' => 'rtf', 'HTML' => 'html', 'PDF' => 'pdf');
        $path = 'D:\\test\\';
        $source = $path.'testwatermark.docx';
        $phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
        foreach($phpWord->getSections() as $section){
            $header = $section->addHeader();
            $header->addWatermark($path.'testwatermark.jpg', array('marginTop' => 0, 'marginLeft' => 0));
        }
        //echo write($phpWord, basename(__FILE__, '.php'), $writers);
        //var_dump($phpWord);exit;
        $phpWord->save($path.'testwatermark-'.time().'.docx', 'Word2007');
        echo 'success';
    }

    // function write($phpWord, $filename, $writers)
    // {
    //     $result = '';
    //     // Write documents
    //     foreach ($writers as $format => $extension) {
    //         $result .= date('H:i:s') . " Write to {$format} format";
    //         if (null !== $extension) {
    //             $targetFile = __DIR__ . "/results/{$filename}.{$extension}";
    //             $phpWord->save($targetFile, $format);
    //         } else {
    //             $result .= ' ... NOT DONE!';
    //         }
    //         $result .= EOL;
    //     }
    //     $result .= getEndingNotes($writers, $filename);
    //     return $result;
    // }
}
