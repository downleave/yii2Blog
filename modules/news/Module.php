<?php

namespace app\modules\news;

/**
 * news module definition class
 */
class Module extends \yii\base\Module
{
    /**
     * {@inheritdoc}
     */
    public $controllerNamespace = 'app\modules\news\controllers';

    /**
     * {@inheritdoc}
     */
    public function init()
    {
        parent::init();

        // custom initialization code goes here
    }

    public function bootstrap($app)
    {
        $app->getUrlManager()->addRules([
            [
                'class' => 'yii\web\UrlRule',
                'route' => $this->id . '/<controller>/<action>',
                'pattern' => $this->id . '/<controller:[\w\-]+>/<action:[\w\-]+>',
            ]
        ], false);
    }
}
