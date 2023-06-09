<?php
/**
Template Name: 赞助我们
 **/
get_header(); ?>

    <body mpa-version="7.16.12" mpa-extension-id="aidjohbjielfdhcaookdaolppglahebo" data-new-gr-c-s-check-loaded="14.990.0"
          data-gr-ext-installed="">
<!--header-->
<main id="body">
    <div class="container">
        <div class="divider"></div>

        <div class="post-body">
            <!--首页导航 需要自己添加进去-->

            <!--内容-->
            <div class="card">
                <div class="card-header">赞助我们</div>
                <div class="card-body ajax-body">
                    <div class="row">
                        <div class="col-5">
                            <div class="nav flex-column nav-pills text-center" id="v-pills-tab" role="tablist" aria-orientation="vertical">
                                <a class="nav-link active show" id="v-pills-home-tab" data-toggle="pill" href="#v-pills-home" role="tab" aria-controls="v-pills-home" aria-selected="true">支付宝</a>
                                <a class="nav-link" id="v-pills-profile-tab" data-toggle="pill" href="#v-pills-profile" role="tab" aria-controls="v-pills-profile" aria-selected="false">微信</a>
                            </div>
                        </div>
                        <div class="col-7">
                            <div class="tab-content" id="v-pills-tabContent">
                                <div class="tab-pane fade active show" id="v-pills-home" role="tabpanel" aria-labelledby="v-pills-home-tab"><img src="<?php
                                    $options = get_option( 'baolog_framework' );
                                    echo $options['baolog-support-alipay'];
                                    ?>"></div>
                                <div class="tab-pane fade" id="v-pills-profile" role="tabpanel" aria-labelledby="v-pills-profile-tab"><img src="<?php
                                    $options = get_option( 'baolog_framework' );
                                    echo $options['baolog-support-wechat'];
                                    ?>"></div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

        </div>

    </div>
</main>

<!-- sidebar -->
<?php
get_sidebar();
?>
<!--footer-->
<?php get_footer(); ?>