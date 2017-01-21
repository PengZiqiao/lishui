from flask import Flask, render_template
from flask_nav import Nav
from flask_nav.elements import Navbar, View
from flask_bootstrap import Bootstrap

nav = Nav()


@nav.navigation()
def navbar():
    return Navbar(
        '溧水月报',
        View('市场运行现状', 'xian_zhuang'),
        View('镇街对比', 'zhen_jie'),
        View('区域对比', 'qu_yu'),
        View('图表部分', 'tu_biao'),
    )


app = Flask(__name__)
nav.init_app(app)
Bootstrap(app)


@app.route('/')
def index():
    """首页"""
    return render_template('index.html')


@app.route('/xianzhuang')
def xian_zhuang():
    """一、市场运行现状
        1、商品房供销情况
    """
    from models.xianzhuang import xianzhuang_data
    data = xianzhuang_data()
    return render_template('xianzhuang.html', data=data)


@app.route('/zhenjie')
def zhen_jie():
    """二、镇街对比"""
    from models.zhenjie import zhenjie_data
    data = zhenjie_data()
    return render_template('zhenjie.html', data=data)


@app.route('/quyu')
def qu_yu():
    """三、区域对比"""
    return render_template('quyu.html')


@app.route('/tubiao')
def tu_biao():
    """五、图表部分"""
    from models.xianzhuang import xianzhuang_data
    data = xianzhuang_data()
    for period in ['当月', '累计']:
        for sheet in data[period].keys():
            for key in data[period][sheet].keys():
                data[period][sheet][key][1] = data[period][sheet][key][1].replace('增长', '').replace('下降', '-').replace(
                    '%', "")
                data[period][sheet][key][2] = data[period][sheet][key][1].replace('增长', '').replace('下降', '-').replace(
                    '%', "")
    dangyue = data['当月']
    leiji = data['累计']
    return render_template('tubiao.html', dangyue=dangyue, leiji=leiji)


if __name__ == '__main__':
    app.run(port=8000)
