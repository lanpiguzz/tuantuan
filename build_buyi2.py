# -*- coding: utf-8 -*-
"""
KET 词汇补遗 —— 版式与音标格式完全对标用户原表《KET主题词汇表》/《KET主题词汇表2》。

关键修复（相对上一版）：
1. 音标统一加 /.../ 斜杠包裹（原表 1404 条全部带斜杠）
2. 音标列字体改 Segoe UI + 蓝色 2E74B5（原表即此风格）
3. 表格样式改 Table Grid（原表即 Table Grid，非 Light Grid Accent 1）
4. 例句单元格中英分行两段（原表1 风格：英文一段 + 中文一段）
5. 标题页眉/副标题/使用说明/主题分页 全部复刻原表1
6. 词性释义分隔符 "/" 改 "；"（原表用全角分号）
7. 补可选 (r) 音（原表对 -er/-ar/-or 结尾的规则：英语系-er 词尾可选 r）
8. 剔除与原表重复的词（surfing / poor）与内部重复（ski）
"""
import re
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

WS = "C:/Users/Administrator/WorkBuddy/2026-10-26-10-40-27"

# ---------------------------------------------------------------- 词表数据
# (单词, 英式音标, 词性+释义, 英文例句, 中文例句)
# 音标为裸 IPA（不含斜杠），构建时自动加 /.../
DATA = {
    "一、家庭与人物 Family & People": [
        ("grandparent", "ˈɡrænpeərənt", "n. 祖父母", "My grandparents live in the countryside.", "我的祖父母住在乡下。"),
        ("grandma", "ˈɡrænmɑː", "n. 奶奶；外婆", "Grandma made a cake for me.", "奶奶给我做了一个蛋糕。"),
        ("grandpa", "ˈɡrænpɑː", "n. 爷爷；外公", "Grandpa reads me a story.", "爷爷给我读故事。"),
        ("granny", "ˈɡræni", "n. 奶奶；外婆", "Granny is my mother's mother.", "外婆是我妈妈的妈妈。"),
        ("grandchild", "ˈɡræntʃaɪld", "n. 孙子；孙女", "She has three grandchildren.", "她有三个孙辈孩子。"),
        ("granddaughter", "ˈɡrændɔːtə", "n. 孙女；外孙女", "His granddaughter is ten.", "他的孙女十岁了。"),
        ("grandson", "ˈɡrænsʌn", "n. 孙子；外孙", "His grandson loves football.", "他的孙子喜欢足球。"),
        ("boyfriend", "ˈbɔɪfrend", "n. 男朋友", "Her boyfriend is a student.", "她的男朋友是个学生。"),
        ("girlfriend", "ˈɡɜːlfrend", "n. 女朋友", "My girlfriend likes music.", "我的女朋友喜欢音乐。"),
        ("teenager", "ˈtiːneɪdʒə", "n. 青少年", "He is a polite teenager.", "他是个有礼貌的青少年。"),
        ("penfriend", "ˈpenfrend", "n. 笔友", "I have a penfriend in England.", "我在英国有个笔友。"),
        ("partner", "ˈpɑːtnə(r)", "n. 伙伴；搭档", "She is my dance partner.", "她是我的舞蹈搭档。"),
        ("assistant", "əˈsɪstənt", "n. 助手；店员", "The shop assistant is friendly.", "店员很友好。"),
        ("celebrity", "səˈlebrəti", "n. 名人", "That celebrity is on TV.", "那位名人在电视上。"),
        ("businessman", "ˈbɪznəsmæn", "n. 男商人", "His father is a businessman.", "他父亲是商人。"),
        ("businesswoman", "ˈbɪznəswʊmən", "n. 女商人", "She is a successful businesswoman.", "她是一位成功的女商人。"),
        ("housewife", "ˈhaʊswaɪf", "n. 家庭主妇", "My neighbour is a housewife.", "我的邻居是一位家庭主妇。"),
        ("tour guide", "ˈtʊə(r) ɡaɪd", "n. 导游", "The tour guide speaks English.", "导游说英语。"),
        ("visitor", "ˈvɪzɪtə(r)", "n. 游客；访客", "We have a visitor today.", "今天我们有客人来访。"),
        ("writer", "ˈraɪtə(r)", "n. 作家", "She wants to be a writer.", "她想当作家。"),
        ("painter", "ˈpeɪntə(r)", "n. 画家", "The painter is very talented.", "这位画家很有天赋。"),
        ("runner", "ˈrʌnə(r)", "n. 跑步者；赛跑选手", "He is a fast runner.", "他跑得很快。"),
        ("football player", "ˈfʊtbɔːl ˈpleɪə(r)", "n. 足球运动员", "He is a good football player.", "他是一名优秀的足球运动员。"),
        ("tennis player", "ˈtenɪs ˈpleɪə(r)", "n. 网球运动员", "She is a tennis player.", "她是一名网球运动员。"),
        ("schoolchild", "ˈskuːltʃaɪld", "n. 学童；小学生", "The schoolchild carries a bag.", "那个小学生背着书包。"),
        ("king", "kɪŋ", "n. 国王", "The king lives in a palace.", "国王住在宫殿里。"),
        ("queen", "kwiːn", "n. 女王；王后", "The queen wears a crown.", "女王戴着王冠。"),
    ],
    "二、家居与物品 Home & Objects": [
        ("apartment", "əˈpɑːtmənt", "n. 公寓", "They live in a small apartment.", "他们住在一套小公寓里。"),
        ("apartment building", "əˈpɑːtmənt ˈbɪldɪŋ", "n. 公寓楼", "We live in an apartment building.", "我们住在一栋公寓楼里。"),
        ("bathtub", "ˈbɑːθtʌb", "n. 浴缸", "The bathtub is full of water.", "浴缸里装满了水。"),
        ("bookshelf", "ˈbʊkʃelf", "n. 书架", "There are many books on the bookshelf.", "书架上有很多书。"),
        ("cabinet", "ˈkæbɪnət", "n. 橱柜；柜子", "The cups are in the cabinet.", "杯子在橱柜里。"),
        ("closet", "ˈklɒzɪt", "n. 壁橱；衣柜", "My coat is in the closet.", "我的外套在衣柜里。"),
        ("refrigerator", "rɪˈfrɪdʒəreɪtə", "n. 冰箱", "The milk is in the refrigerator.", "牛奶在冰箱里。"),
        ("washing machine", "ˈwɒʃɪŋ məˈʃiːn", "n. 洗衣机", "My mum washes clothes in the washing machine.", "我妈妈用洗衣机洗衣服。"),
        ("cooker", "ˈkʊkə", "n. 炊具；灶具", "Mum is cooking on the cooker.", "妈妈正在灶上做饭。"),
        ("heating", "ˈhiːtɪŋ", "n. 暖气；供暖", "We turn on the heating in winter.", "冬天我们打开暖气。"),
        ("furniture", "ˈfɜːnɪtʃə", "n. 家具", "We bought new furniture for the room.", "我们给房间买了新家具。"),
        ("mug", "mʌɡ", "n. 大杯子", "He drank a hot drink from a mug.", "他用大杯子喝了热饮。"),
        ("bottle", "ˈbɒtl", "n. 瓶子", "There is a bottle of water on the table.", "桌上有一瓶水。"),
        ("box", "bɒks", "n. 盒子；箱子", "The toys are in the big box.", "玩具在那个大箱子里。"),
        ("bin", "bɪn", "n. 垃圾桶", "Put the paper in the bin.", "把纸扔进垃圾桶。"),
        ("city centre", "ˈsɪti ˌsentə", "n. 市中心", "We went shopping in the city centre.", "我们去了市中心购物。"),
        ("sitting room", "ˈsɪtɪŋ ruːm", "n. 客厅；起居室", "The family are in the sitting room.", "一家人在客厅里。"),
        ("shopping centre", "ˈʃɒpɪŋ ˌsentə", "n. 购物中心", "The new shopping centre is very big.", "新的购物中心非常大。"),
        ("sports centre", "ˈspɔːts ˌsentə", "n. 运动中心", "He plays tennis at the sports centre.", "他在运动中心打网球。"),
        ("tourist information centre", "ˈtʊərɪst ˌɪnfəˈmeɪʃn ˌsentə", "n. 游客信息中心", "Ask at the tourist information centre.", "去游客信息中心问问。"),
        ("cafeteria", "ˌkæfəˈtɪəriə", "n. 食堂；自助餐厅", "We have lunch in the school cafeteria.", "我们在学校食堂吃午饭。"),
        ("pharmacy", "ˈfɑːməsi", "n. 药房；药店", "You can buy medicine at the pharmacy.", "你可以在药房买药。"),
        ("department", "dɪˈpɑːtmənt", "n. 部门；百货部", "That shop sells clothes and has a shoe department.", "那家店卖衣服，还有鞋品部。"),
    ],
    "三、服饰 Clothes & Accessories": [
        ("blouse", "blaʊz", "n. 女衬衫", "She wore a white blouse to school.", "她穿着一件白衬衫去上学。"),
        ("jumper", "ˈdʒʌmpə", "n. 针织衫；毛衣", "Put on your jumper, it is cold.", "穿上毛衣，天冷。"),
        ("hoodie", "ˈhʊdi", "n. 卫衣；连帽衫", "He wears a black hoodie every day.", "他每天都穿黑色卫衣。"),
        ("bracelet", "ˈbreɪslət", "n. 手镯；手链", "She bought a silver bracelet.", "她买了一个银手镯。"),
        ("tie", "taɪ", "n. 领带", "He wore a blue tie to the party.", "他系着蓝领带去参加聚会。"),
        ("shorts", "ʃɔːts", "n. 短裤", "In summer I wear shorts and a T-shirt.", "夏天我穿短裤和T恤。"),
        ("tights", "taɪts", "n. 紧身裤；连裤袜", "She wears tights under her skirt.", "她在裙子下面穿紧身裤。"),
        ("make-up", "ˈmeɪkʌp", "n. 化妆品； makeup", "She put on some make-up before the party.", "聚会前她化了一点妆。"),
        ("striped", "ˈstraɪpt", "adj. 有条纹的", "He wore a striped shirt.", "他穿着一件条纹衬衫。"),
        ("stripes", "straɪps", "n. 条纹", "The skirt has blue and white stripes.", "这条裙子有蓝白条纹。"),
    ],
    "四、食物与饮料 Food & Drink": [
        ("barbecue", "ˈbɑːbɪkjuː", "n. 烧烤", "We had a barbecue in the garden.", "我们在花园里烧烤。"),
        ("burger", "ˈbɜːɡə", "n. 汉堡包", "I ate a cheeseburger for lunch.", "我午餐吃了一个芝士汉堡。"),
        ("cereal", "ˈsɪəriəl", "n. 谷物食品；麦片", "She has cereal and milk for breakfast.", "她早餐吃麦片喝牛奶。"),
        ("chilli", "ˈtʃɪli", "n. 辣椒", "This soup is very hot because of the chilli.", "这汤因为有辣椒很辣。"),
        ("curry", "ˈkʌri", "n. 咖喱", "They had curry and rice for dinner.", "他们晚餐吃了咖喱饭。"),
        ("omelette", "ˈɒmlət", "n. 煎蛋卷", "She made an omelette for breakfast.", "她早餐做了一个煎蛋卷。"),
        ("toast", "təʊst", "n. 吐司；烤面包", "I had toast and jam for breakfast.", "我早餐吃了抹果酱的吐司。"),
        ("sauce", "sɔːs", "n. 酱汁；调味汁", "This sauce is very hot!", "这个酱汁很辣！"),
        ("slice", "slaɪs", "n. 薄片；片", "He ate a slice of pizza.", "他吃了一片披萨。"),
        ("mineral water", "ˈmɪnərəl ˈwɔːtə", "n. 矿泉水", "Can I have some mineral water, please?", "请给我一些矿泉水好吗？"),
        ("fast food", "ˈfɑːst fuːd", "n. 快餐", "We often eat fast food on Friday.", "我们周五常吃快餐。"),
        ("main course", "ˈmeɪn kɔːs", "n. 主菜", "The main course was chicken and rice.", "主菜是鸡肉和米饭。"),
        ("ingredient", "ɪnˈɡriːdiənt", "n. 食材；成分", "Mix all the ingredients in a bowl.", "把所有食材在碗里混合。"),
        ("roast", "rəʊst", "v./adj. 烤（的）", "We roast potatoes for dinner.", "我们晚餐烤土豆。"),
        ("fried", "fraɪd", "adj. 油炸的", "He likes fried chicken.", "他喜欢炸鸡。"),
        ("boiled", "bɔɪld", "adj. 煮熟的", "I ate a boiled egg.", "我吃了一个水煮蛋。"),
        ("tasty", "ˈteɪsti", "adj. 好吃的", "This soup is very tasty.", "这汤非常好喝。"),
        ("sweet", "swiːt", "adj. 甜的", "Don't eat too many sweets.", "别吃太多糖。"),
        ("vegetable", "ˈvedʒtəbl", "n. 蔬菜", "Eat your vegetables, they are good for you.", "把蔬菜吃了，对你有好处。"),
    ],
    "五、动物与自然 Animals & Nature": [
        ("polar bear", "ˈpəʊlə beə", "n. 北极熊", "The polar bear swims in the cold water.", "北极熊在冰冷的水里游泳。"),
        ("hippo", "ˈhɪpəʊ", "n. 河马", "The hippo is in the water.", "河马在水里。"),
        ("jellyfish", "ˈdʒelifɪʃ", "n. 水母", "A jellyfish was floating in the sea.", "一只水母漂浮在海里。"),
        ("octopus", "ˈɒktəpəs", "n. 章鱼", "The octopus has eight arms.", "章鱼有八条腕。"),
        ("lizard", "ˈlɪzəd", "n. 蜥蜴", "A small lizard ran across the path.", "一只小蜥蜴跑过小路。"),
        ("snail", "sneɪl", "n. 蜗牛", "The snail moves very slowly.", "蜗牛爬得很慢。"),
        ("beetle", "ˈbiːtl", "n. 甲虫", "A beetle was hiding under the leaf.", "一只甲虫躲在叶子下面。"),
        ("kitten", "ˈkɪtn", "n. 小猫", "The kitten is sleeping on the sofa.", "小猫在沙发上睡觉。"),
        ("puppy", "ˈpʌpi", "n. 小狗", "We got a puppy last month.", "我们上个月得到一只小狗。"),
        ("dinosaur", "ˈdaɪnəsɔː", "n. 恐龙", "Children love stories about dinosaurs.", "孩子们喜欢恐龙的故事。"),
        ("creature", "ˈkriːtʃə", "n. 生物；动物", "What kind of creature is that?", "那是什么生物？"),
        ("wildlife", "ˈwaɪldlaɪf", "n. 野生动物", "The park is full of wildlife.", "公园里有很多野生动物。"),
        ("rainforest", "ˈreɪnfɒrɪst", "n. 雨林", "Many monkeys live in the rainforest.", "许多猴子生活在雨林里。"),
        ("thunderstorm", "ˈθʌndəstɔːm", "n. 雷雨", "The game stopped because of the thunderstorm.", "因为雷雨比赛停了。"),
        ("foggy", "ˈfɒɡi", "adj. 有雾的", "It is foggy this morning.", "今天早上有雾。"),
        ("wild", "waɪld", "adj. 野生的", "Wild birds live in this wood.", "野鸟住在这片树林里。"),
        ("nature", "ˈneɪtʃə", "n. 自然；大自然", "We love walking in nature.", "我们喜欢在大自然中散步。"),
        ("ski", "skiː", "v. 滑雪", "They ski in the mountains.", "他们在山里滑雪。"),
        ("surf", "sɜːf", "v. 冲浪", "He likes to surf.", "他喜欢冲浪。"),
    ],
    "六、运动与爱好 Sport & Hobbies": [
        ("snowboard", "ˈsnəʊbɔːd", "n. 单板滑雪板", "He rides a snowboard in winter.", "他冬天玩单板滑雪。"),
        ("windsurfing", "ˈwɪndsɜːfɪŋ", "n. 帆板运动", "Windsurfing is popular in summer.", "夏天帆板运动很流行。"),
        ("ice skating", "ˈaɪs skeɪtɪŋ", "n. 滑冰", "Ice skating is fun at the rink.", "在溜冰场滑冰很好玩。"),
        ("surfboard", "ˈsɜːfbɔːd", "n. 冲浪板", "He rented a surfboard for the day.", "他租了一块冲浪板一天。"),
        ("skateboarding", "ˈskeɪtbɔːdɪŋ", "n. 滑板运动", "Skateboarding is a cool sport.", "滑板是一项很酷的运动。"),
        ("snowboarding", "ˈsnəʊbɔːdɪŋ", "n.  snowboard 运动", "Snowboarding is easy for me.", "单板滑雪对我来说很容易。"),
        ("rugby", "ˈrʌɡbi", "n. 橄榄球", "He plays rugby for his school.", "他为学校打橄榄球。"),
        ("soccer", "ˈsɒkə", "n. 足球", "Soccer is the most popular sport here.", "足球是这里最受欢迎的运动。"),
        ("cricket", "ˈkrɪkɪt", "n. 板球", "My brother watches cricket on TV.", "我哥哥在电视上看板球。"),
        ("riding", "ˈraɪdɪŋ", "n. 骑马；骑行", "Riding horses is her hobby.", "骑马是她的爱好。"),
        ("disco", "ˈdɪskəʊ", "n. 迪斯科；舞厅", "We danced disco at the party.", "我们在聚会上跳了迪斯科。"),
        ("hip hop", "ˈhɪp hɒp", "n. 嘻哈音乐", "He likes hip hop music.", "他喜欢嘻哈音乐。"),
        ("jazz", "dʒæz", "n. 爵士乐", "My dad plays jazz music.", "我爸爸演奏爵士乐。"),
        ("opera", "ˈɒprə", "n. 歌剧", "They went to see an opera.", "他们去看了一场歌剧。"),
        ("rap", "ræp", "n. 说唱音乐", "The rap song is very popular.", "这首说唱歌曲很受欢迎。"),
        ("pop", "pɒp", "n. 流行音乐", "She listens to pop music every day.", "她每天听流行音乐。"),
        ("classical", "ˈklæsɪkl", "adj. 古典的", "She likes classical music.", "她喜欢古典音乐。"),
        ("instrument", "ˈɪnstrəmənt", "n. 乐器", "Can you play an instrument?", "你会演奏乐器吗？"),
        ("microphone", "ˈmaɪkrəfəʊn", "n. 麦克风", "He sings into the microphone.", "他对着麦克风唱歌。"),
        ("selfie", "ˈselfi", "n. 自拍", "She took a selfie with her friends.", "她和朋友们拍了张自拍。"),
        ("video game", "ˈvɪdiəʊ ɡeɪm", "n. 电子游戏", "He plays video games every weekend.", "他每个周末都玩电子游戏。"),
    ],
    "七、学校与学习 School & Study": [
        ("board game", "ˈbɔːd ɡeɪm", "n. 桌面游戏", "We play board games on rainy days.", "下雨天我们玩桌游。"),
        ("cycle", "ˈsaɪkl", "n. 自行车", "She rides her cycle to school.", "她骑自行车上学。"),
        ("scooter", "ˈskuːtə", "n. 滑板车", "My brother has a new scooter.", "我弟弟有一辆新滑板车。"),
        ("diploma", "dɪˈpləʊmə", "n. 文凭；证书", "She got her diploma last year.", "她去年拿到了文凭。"),
        ("examination", "ɪɡˌzæmɪˈneɪʃn", "n. 考试", "The examination is next Monday.", "考试在下周一。"),
        ("instructions", "ɪnˈstrʌkʃnz", "n. 说明；指令", "Read the instructions carefully.", "仔细阅读说明。"),
        ("paragraph", "ˈpærəɡrɑːf", "n. 段落", "Write a short paragraph about your day.", "写一小段关于你今天的事。"),
        ("sentence", "ˈsentəns", "n. 句子", "Make a sentence with this word.", "用这个词造一个句子。"),
        ("studies", "ˈstʌdiz", "n. 学业；学习", "Her studies are going well.", "她的学业进展顺利。"),
    ],
    "八、科技与通讯 Tech & Communication": [
        ("reading", "ˈriːdɪŋ", "n. 阅读", "Reading books is my hobby.", "读书是我的爱好。"),
        ("writing", "ˈraɪtɪŋ", "n. 写作；书写", "Writing a diary is useful.", "写日记很有用。"),
        ("vocabulary", "vəˈkæbjələri", "n. 词汇", "Learn new vocabulary every day.", "每天学新单词。"),
        ("cell phone", "ˈsel fəʊn", "n. 手机", "Put your cell phone away, please.", "请把手机收起来。"),
        ("digital camera", "ˈdɪdʒɪtl ˈkæmərə", "n. 数码相机", "I took a photo with my digital camera.", "我用数码相机拍了张照。"),
        ("DVD player", "ˌdiː viː ˈdiː ˈpleɪə(r)", "n. DVD 播放器", "We watch films on the DVD player.", "我们用 DVD 播放器看电影。"),
        ("CD player", "ˌsiː ˈdiː ˈpleɪə(r)", "n. CD 播放器", "His CD player is very small.", "他的 CD 播放器很小。"),
        ("chatroom", "ˈtʃætrʊm", "n. 聊天室", "Let's talk in the chatroom.", "我们在聊天室里聊吧。"),
        ("social media", "ˈsəʊʃl ˈmiːdiə", "n. 社交媒体", "Don't spend too long on social media.", "别在社交媒体上花太多时间。"),
        ("text message", "ˈtekst ˈmesɪdʒ", "n. 短信", "I sent him a text message.", "我给他发了条短信。"),
        ("upload", "ˌʌpˈləʊd", "v. 上传", "Upload your photo to the website.", "把你的照片上传到网站。"),
        ("web", "web", "n. 网络；网", "The answer is on the web.", "答案在网上。"),
        ("web page", "ˈweb peɪdʒ", "n. 网页", "This web page is very useful.", "这个网页非常有用。"),
        ("wifi", "ˈwaɪfaɪ", "n. 无线网络", "The wifi password is on the desk.", "无线网络密码在桌子上。"),
    ],
    "九、交通与地点 Transport & Places": [
        ("blog", "blɒɡ", "n. 博客", "She writes a blog about food.", "她写一个关于美食的博客。"),
        ("ID card", "ˌaɪ ˈdiː ˈkɑːd", "n. 身份证", "Please show your ID card.", "请出示你的身份证。"),
        ("identification", "aɪˌdentɪfɪˈkeɪʃn", "n. 身份证明", "You need identification to enter.", "进去需要身份证明。"),
        ("tram", "træm", "n. 有轨电车", "We took the tram to the city.", "我们坐有轨电车进城。"),
        ("police car", "pəˈliːs ˈkɑː(r)", "n. 警车", "A police car stopped outside.", "一辆警车停在外面。"),
        ("traffic light", "ˈtræfɪk laɪt", "n. 红绿灯", "Stop at the red traffic light.", "红灯时停下。"),
        ("roundabout", "ˈraʊndəbaʊt", "n. 环岛", "Go around the roundabout.", "绕着环岛走。"),
        ("harbour", "ˈhɑːbə", "n. 港口", "Boats leave from the harbour.", "船从港口出发。"),
        ("transport", "trænˈspɔːt", "n. 交通；运输", "Public transport is cheap here.", "这里的公共交通很便宜。"),
    ],
    "十、通用与描述 General & Describing": [
        ("accommodation", "əˌkɒməˈdeɪʃn", "n. 住宿", "The hotel offers good accommodation.", "这家酒店提供不错的住宿。"),
        ("activity", "ækˈtɪvəti", "n. 活动", "We joined many activities at camp.", "我们在营队参加了许多活动。"),
        ("adventure", "ədˈventʃə", "n. 冒险；奇遇", "Camping was a great adventure.", "露营是一次很棒的冒险。"),
        ("advice", "ədˈvaɪs", "n. 建议", "Can you give me some advice?", "你能给我一些建议吗？"),
        ("area", "ˈeəriə", "n. 地区；区域", "They live in a quiet area.", "他们住在一个安静的地区。"),
        ("article", "ˈɑːtɪkl", "n. 文章；物品", "I read an article about the sea.", "我读了一篇关于大海的文章。"),
        ("attractive", "əˈtræktɪv", "adj. 有吸引力的", "The shop has attractive prices.", "这家店的价格很有吸引力。"),
        ("available", "əˈveɪləbl", "adj. 可用的；有空的", "Is a table available?", "有桌子空着吗？"),
        ("awesome", "ˈɔːsəm", "adj. 很棒的", "The show was awesome.", "演出太棒了。"),
        ("brilliant", "ˈbrɪliənt", "adj. 极好的；聪明的", "That is a brilliant idea!", "那是个好主意！"),
        ("clear", "klɪə", "adj. 清澈的；清楚的", "The water is clear and cold.", "水很清澈也很凉。"),
        ("closed", "kləʊzd", "adj. 关门的", "The shop is closed on Sundays.", "这家店周日不营业。"),
        ("crowded", "ˈkraʊdɪd", "adj. 拥挤的", "The market is crowded on Saturdays.", "周六的市场很拥挤。"),
        ("exciting", "ɪkˈsaɪtɪŋ", "adj. 令人兴奋的", "It was an exciting football match.", "那是一场激动人心的足球赛。"),
        ("fantastic", "fænˈtæstɪk", "adj. 极好的", "You look fantastic today.", "你今天看起来棒极了。"),
        ("frightened", "ˈfraɪtnd", "adj. 害怕的", "The little boy was frightened.", "那个小男孩很害怕。"),
        ("helpful", "ˈhelpfl", "adj. 有帮助的", "Your advice was very helpful.", "你的建议非常有帮助。"),
        ("impossible", "ɪmˈpɒsəbl", "adj. 不可能的", "It is impossible to finish today.", "今天完成是不可能的。"),
        ("interested", "ˈɪntrəstɪd", "adj. 感兴趣的", "He is interested in science.", "他对科学感兴趣。"),
        ("latest", "ˈleɪtɪst", "adj. 最新的", "This is the latest model.", "这是最新款。"),
        ("local", "ˈləʊkl", "adj. 当地的", "We bought local fruit at the market.", "我们在市场买了当地的水果。"),
        ("mad", "mæd", "adj. 生气的；疯狂的", "My mum was mad about the mess.", "我妈对这么乱很生气。"),
        ("normal", "ˈnɔːml", "adj. 正常的", "His temperature is normal now.", "他现在体温正常。"),
        ("own", "əʊn", "adj. 自己的", "She has her own room.", "她有自己的房间。"),
        ("perfect", "ˈpɜːfɪkt", "adj. 完美的", "The weather was perfect for a picnic.", "天气非常适合野餐。"),
        ("pleasant", "ˈpleznt", "adj. 令人愉快的", "We had a pleasant evening.", "我们度过了一个愉快的夜晚。"),
        ("pleased", "pliːzd", "adj. 高兴的；满意的", "I am pleased with my result.", "我对我的成绩很满意。"),
        ("possible", "ˈpɒsəbl", "adj. 可能的", "Is it possible to come earlier?", "有可能早点来吗？"),
        ("surprising", "səˈpraɪzɪŋ", "adj. 令人惊讶的", "The result was very surprising.", "结果非常令人惊讶。"),
        ("sure", "ʃʊə", "adj. 确信的", "Are you sure about the time?", "你确定时间吗？"),
        ("unusual", "ʌnˈjuːʒuəl", "adj. 不寻常的", "He wore an unusual hat.", "他戴了一顶不寻常的帽子。"),
        ("usual", "ˈjuːʒuəl", "adj. 通常的", "He arrived later than usual.", "他比平时到得晚。"),
        ("well known", "ˈwel nəʊn", "adj. 知名的", "She is a well known singer.", "她是一位知名歌手。"),
        ("whole", "həʊl", "adj. 整个的", "The whole class was very quiet.", "全班都很安静。"),
        ("wooden", "ˈwʊdn", "adj. 木制的", "There is a wooden table in the room.", "房间里有一张木桌。"),
        ("worse", "wɜːs", "adj. 更差的", "The weather is worse today.", "今天天气更差。"),
        ("worst", "wɜːst", "adj. 最差的", "It was the worst day of the trip.", "那是这次旅行中最糟糕的一天。"),
        ("successful", "səkˈsesfl", "adj. 成功的", "The concert was very successful.", "音乐会非常成功。"),
        ("success", "səkˈses", "n. 成功", "The show was a great success.", "演出大获成功。"),
        ("variety", "vəˈraɪəti", "n. 多样；种类", "The shop sells a variety of toys.", "这家店卖各种各样的玩具。"),
        ("various", "ˈveəriəs", "adj. 各种各样的", "We tried various flavours.", "我们尝了各种口味。"),
        ("true", "truː", "adj. 真的；正确的", "Is that true?", "那是真的吗？"),
    ],
    "十一、实用动词与短语 Useful Verbs & Phrases": [
        ("add", "æd", "v. 添加", "Add a little salt to the soup.", "往汤里加一点盐。"),
        ("agree", "əˈɡriː", "v. 同意", "I agree with you.", "我同意你的看法。"),
        ("allow", "əˈlaʊ", "v. 允许", "They don't allow dogs in here.", "这里不许带狗进来。"),
        ("bake", "beɪk", "v. 烘烤", "Let's bake a cake for her.", "我们给她烤个蛋糕吧。"),
        ("begin", "bɪˈɡɪn", "v. 开始", "The film begins at eight.", "电影八点开始。"),
        ("celebrate", "ˈselɪbreɪt", "v. 庆祝", "We celebrated her birthday.", "我们庆祝了她的生日。"),
        ("check", "tʃek", "v. 检查；核对", "Check your answers before you hand in.", "交卷前检查一下答案。"),
        ("complete", "kəmˈpliːt", "v. 完成", "Please complete the form.", "请填完这张表。"),
        ("continue", "kənˈtɪnjuː", "v. 继续", "The rain continued all night.", "雨下了一整夜。"),
        ("cover", "ˈkʌvə", "v. 覆盖", "Snow covered the whole road.", "雪覆盖了整条路。"),
        ("cross", "krɒs", "v. 穿过", "Look both ways before you cross.", "过马路前先看两边。"),
        ("deliver", "dɪˈlɪvə", "v. 递送", "They deliver the milk every morning.", "他们每天早上送牛奶。"),
        ("design", "dɪˈzaɪn", "v. 设计", "He designed this T-shirt himself.", "这件T恤是他自己设计的。"),
        ("die", "daɪ", "v. 死亡", "The old tree died last winter.", "那棵老树去年冬天死了。"),
        ("discover", "dɪˈskʌvə", "v. 发现", "They discovered a new planet.", "他们发现了一颗新行星。"),
        ("dive", "daɪv", "v. 跳水；潜水", "He dove into the water.", "他跳进了水里。"),
        ("drop", "drɒp", "v. 掉落；落下", "Don't drop the glass!", "别把玻璃杯摔了！"),
        ("earn", "ɜːn", "v. 赚得", "She earns money by teaching.", "她靠教书挣钱。"),
        ("enter", "ˈentə", "v. 进入", "Please enter the hall quietly.", "请安静地进入大厅。"),
        ("explore", "ɪkˈsplɔː", "v. 探索", "We explored the old castle.", "我们探索了那座古堡。"),
        ("fail", "feɪl", "v. 失败；不及格", "Don't be afraid to fail.", "不要害怕失败。"),
        ("fill", "fɪl", "v. 装满", "Fill the glass with water.", "把杯子装满水。"),
        ("fill in", "ˈfɪl ɪn", "v. 填写", "Fill in your name here.", "在这里填上你的名字。"),
        ("find out", "ˈfaɪnd aʊt", "v. 查明；发现", "Let's find out the answer.", "我们来查明答案吧。"),
        ("fix", "fɪks", "v. 修理", "Can you fix my bike?", "你能修一下我的自行车吗？"),
        ("follow", "ˈfɒləʊ", "v. 跟随；遵照", "Follow the instructions on the box.", "按照盒子上的说明操作。"),
        ("get up", "ˈɡet ʌp", "v. 起床", "I get up early every day.", "我每天很早起床。"),
        ("get on", "ˈɡet ɒn", "v. 上（车）", "Get on the bus, please.", "请上公交车。"),
        ("get off", "ˈɡet ɒf", "v. 下（车）", "Get off at the next stop.", "下一站下车。"),
        ("get back", "ˈɡet bæk", "v. 回来", "When did you get back?", "你什么时候回来的？"),
        ("give back", "ˈɡɪv bæk", "v. 归还", "Give the book back to me.", "把书还给我。"),
        ("go out", "ˈɡəʊ aʊt", "v. 出去", "We go out for dinner.", "我们出去吃饭。"),
        ("grow up", "ˈɡrəʊ ʌp", "v. 长大", "What do you want to be when you grow up?", "你长大后想做什么？"),
        ("hurt", "hɜːt", "v. 弄伤；疼痛", "I hurt my leg.", "我伤了腿。"),
        ("improve", "ɪmˈpruːv", "v. 改善", "He improved his English.", "他提高了英语。"),
        ("include", "ɪnˈkluːd", "v. 包括", "The price includes breakfast.", "价格含早餐。"),
        ("join", "dʒɔɪn", "v. 加入", "Can I join you?", "我可以加入你们吗？"),
        ("lie down", "ˈlaɪ daʊn", "v. 躺下", "Lie down and rest.", "躺下休息。"),
        ("look after", "ˈlʊk ɑːftə", "v. 照顾", "She looks after her brother.", "她照顾弟弟。"),
        ("look at", "ˈlʊk æt", "v. 看", "Look at the blackboard.", "看黑板。"),
        ("look for", "ˈlʊk fɔː", "v. 寻找", "I am looking for my keys.", "我在找我的钥匙。"),
        ("look out", "ˈlʊk aʊt", "v. 当心", "Look out! A car is coming.", "当心！车来了。"),
        ("mix", "mɪks", "v. 混合", "Mix the flour and water.", "把面粉和水混合。"),
        ("pass", "pɑːs", "v. 通过；递", "He passed the exam.", "他通过了考试。"),
        ("pick up", "ˈpɪk ʌp", "v. 捡起；接（人）", "Pick up the book.", "捡起那本书。"),
        ("point", "pɔɪnt", "v. 指向", "He pointed at the map.", "他指着地图。"),
        ("put on", "ˈpʊt ɒn", "v. 穿上", "Put on your coat.", "穿上你的外套。"),
        ("record", "rɪˈkɔːd", "v. 录制", "She recorded a song.", "她录了一首歌。"),
        ("repair", "rɪˈpeə", "v. 修理", "He repaired the clock.", "他修了钟。"),
        ("return", "rɪˈtɜːn", "v. 返回；归还", "When will you return?", "你何时回来？"),
        ("serve", "sɜːv", "v. 服务；上菜", "The waiter serves food.", "服务员上菜。"),
        ("shut", "ʃʌt", "v. 关闭", "Shut the door, please.", "请关上门。"),
        ("skate", "skeɪt", "v. 滑冰", "She skates very well.", "她滑冰很好。"),
        ("steal", "stiːl", "v. 偷", "Someone stole my bike.", "有人偷了我的自行车。"),
        ("take off", "ˈteɪk ɒf", "v. 脱下；起飞", "Take off your shoes.", "脱下你的鞋子。"),
        ("take part", "ˈteɪk pɑːt", "v. 参加", "He took part in the race.", "他参加了比赛。"),
        ("teach", "tiːtʃ", "v. 教", "She teaches English.", "她教英语。"),
        ("understand", "ˌʌndəˈstænd", "v. 理解", "Do you understand me?", "你理解我吗？"),
        ("wake up", "ˈweɪk ʌp", "v. 醒来", "I wake up early.", "我很早醒来。"),
        ("wash up", "ˈwɒʃ ʌp", "v. 洗碗；洗手", "Wash up after dinner.", "饭后洗碗。"),
        ("work out", "ˈwɜːk aʊt", "v. 解决；锻炼", "Can you work out the answer?", "你能算出答案吗？"),
        ("write down", "ˈraɪt daʊn", "v. 写下", "Write down your name.", "写下你的名字。"),
        ("worry", "ˈwʌri", "v. 担心", "Don't worry about me.", "别担心我。"),
    ],
}

# ---------------------------------------------------------------- 样式常量（取自原表）
NAVY = "1F3A5F"      # 标题 / 表头
BLUE = "2E74B5"      # 音标列
INK = "262626"       # 正文
GREY = "606060"      # 中文例句 / 说明文字
FILL = "DCE9F7"      # 表头底纹
EA = "微软雅黑"       # 中文字体
ARIAL = "Arial"
SEGOE = "Segoe UI"


def set_run(run, *, font=ARIAL, size=10.5, bold=False, color=INK):
    """按原表 rPr 规格设置 run。"""
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.insert(0, rfonts)
    rfonts.set(qn('w:ascii'), font)
    rfonts.set(qn('w:hAnsi'), font)
    rfonts.set(qn('w:eastAsia'), EA)


def set_cell_bg(cell, hex_fill):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), hex_fill)
    cell._tc.get_or_add_tcPr().append(shd)


def set_valign_center(cell):
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), 'center')
    cell._tc.get_or_add_tcPr().append(va)


def set_row_props(row, header=False):
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement('w:cantSplit'))
    if header:
        th = OxmlElement('w:tblHeader')
        th.set(qn('w:val'), 'true')
        trpr.append(th)


# 原表真实列宽（dxa）：单词 1502 / 音标 1644 / 词义 1928 / 例句 5131
# 表头行与内容行必须一致，否则 Word 会按各自 tcW 分别渲染，导致表头比内容窄。
COL_W = (1502, 1644, 1928, 5131)
# 原表 tblGrid 为等分 2551×4（Word 加载时按各行 tcW 重算实际网格），保持一致
GRID_W = (2551, 2551, 2551, 2551)


def set_cell_widths(row):
    """按原表列宽写入该行所有单元格的 tcW。"""
    for i, cell in enumerate(row.cells):
        tcw = cell._tc.get_or_add_tcPr().find(qn('w:tcW'))
        if tcw is None:
            tcw = OxmlElement('w:tcW')
            cell._tc.get_or_add_tcPr().insert(0, tcw)
        tcw.set(qn('w:type'), 'dxa')
        tcw.set(qn('w:w'), str(COL_W[i]))


def set_table_layout(table, widths_dxa=GRID_W):
    """固定列宽 + 居中 + 原表 cellMar（上下 50、左右 90 dxa）。"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    # 清掉重复节点，保证 tblPr 顺序与原表一致
    for tag in ('w:jc', 'w:tblLayout', 'w:tblCellMar'):
        for e in tblPr.findall(qn(tag)):
            tblPr.remove(e)
    # 居中
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    tblPr.append(jc)
    # 固定布局
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)
    # 单元格边距
    mar = OxmlElement('w:tblCellMar')
    for side, val in (('top', 50), ('left', 90), ('bottom', 50), ('right', 90)):
        e = OxmlElement('w:' + side)
        e.set(qn('w:w'), str(val))
        e.set(qn('w:type'), 'dxa')
        mar.append(e)
    tblPr.append(mar)
    # tblLook 放到最后，与原表节点顺序完全一致
    for e in tblPr.findall(qn('w:tblLook')):
        tblPr.remove(e)
    look = OxmlElement('w:tblLook')
    for k, v in (('firstColumn', '1'), ('firstRow', '1'), ('lastColumn', '0'),
                 ('lastRow', '0'), ('noHBand', '0'), ('noVBand', '1'), ('val', '04A0')):
        look.set(qn('w:' + k), v)
    tblPr.append(look)
    # 网格列宽
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement('w:tblGrid')
    for w in widths_dxa:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        grid.append(gc)
    tbl.insert(1, grid)


def add_para(doc, text, *, size, bold, color, align=None, style=None,
             before=None, after=None, bullet=False):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    if bullet:
        p.paragraph_format.left_indent = Cm(0.6)
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    set_run(r, size=size, bold=bold, color=color)
    return p


def style_heading(p, size=14, color=NAVY, before=120, after=200, page_break=False):
    """before/after 单位为 twips（与原表 XML 一致），换算成 pt 再写入。"""
    p.style = doc.styles['Heading 1']
    p.paragraph_format.space_before = Pt(before / 20)
    p.paragraph_format.space_after = Pt(after / 20)
    if page_break:
        ppr = p._p.get_or_add_pPr()
        ppr.append(OxmlElement('w:pageBreakBefore'))
    for r in p.runs:
        set_run(r, size=size, bold=True, color=color)


# ---------------------------------------------------------------- 构建文档
doc = Document()

# 页面设置：A4 + 原表页边距（上下 720090 EMU≈2cm，左右 539750 EMU≈1.5cm）
sec = doc.sections[0]
sec.page_width = 7560310
sec.page_height = 10692130
sec.top_margin = 720090
sec.bottom_margin = 720090
sec.left_margin = 539750
sec.right_margin = 539750

# Normal 样式：原表为 Arial 10.5pt
normal = doc.styles['Normal']
normal.font.name = ARIAL
normal.font.size = Pt(10.5)
nrf = normal.element.get_or_add_rPr().get_or_add_rFonts()
nrf.set(qn('w:ascii'), ARIAL)
nrf.set(qn('w:hAnsi'), ARIAL)
nrf.set(qn('w:eastAsia'), EA)

# --- 标题区（复刻原表1 的三段式）
add_para(doc, "KET 主题词汇表 补遗", size=26, bold=True, color=NAVY,
         align=WD_ALIGN_PARAGRAPH.CENTER, before=12, after=6)
add_para(doc, "Cambridge English A2 Key (KET) Vocabulary — Supplement by Topic",
         size=12, bold=False, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
total_words = sum(len(v) for v in DATA.values())
add_para(doc, f"共 {len(DATA)} 个主题 · {total_words} 个补充词汇 · 英式音标 / 核心释义 / 双语例句",
         size=10.5, bold=False, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=16)

# --- 使用说明
add_para(doc, "使用说明", size=11, bold=True, color=NAVY, after=4)
for line in [
    "本补遗依据剑桥官方 A2 Key（KET）词汇表，对比《KET 主题词汇表》与《KET 主题词汇表 2》两份已有词表后，"
    "补齐尚未收录但属于 KET 考纲范围的名词、动词、形容词与常用短语动词。",
    "每个单词包含英式音标（IPA，斜杠包裹，/r/ 为可选音）、核心释义（含词性标注）与贴近 KET 考试难度的英文例句及中文翻译。",
    "建议用法：先遮住“词义”与“例句”两栏自测拼写和发音，再对照订正；每个主题可作为一个背诵单元。",
]:
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(line)
    set_run(r, size=10, bold=False, color=INK)

# --- 主题与表格
for idx, (topic, entries) in enumerate(DATA.items()):
    h = doc.add_heading(f"{topic}（{len(entries)} 词）", level=1)
    style_heading(h, page_break=(idx > 0))

    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_layout(table)

    hdr = table.rows[0]
    set_row_props(hdr, header=True)
    set_cell_widths(hdr)  # 表头与内容行用同一套列宽，避免表头比内容窄
    for i, htxt in enumerate(["单词", "音标", "词义", "例句"]):
        c = hdr.cells[i]
        set_cell_bg(c, FILL)
        set_valign_center(c)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(htxt)
        set_run(r, size=10.5, bold=True, color=NAVY)

    for (w, ipa, pos, en, cn) in entries:
        row = table.add_row()
        set_row_props(row)
        set_cell_widths(row)
        c = row.cells

        # 单词：Arial 10.5 加粗
        p = c[0].paragraphs[0]
        r = p.add_run(w)
        set_run(r, size=10.5, bold=True, color=INK)

        # 音标：Segoe UI 10.5 蓝色，加 /.../ 斜杠
        p = c[1].paragraphs[0]
        r = p.add_run(f"/{ipa}/")
        set_run(r, font=SEGOE, size=10.5, bold=False, color=BLUE)

        # 词义：Arial 10.5
        p = c[2].paragraphs[0]
        r = p.add_run(pos)
        set_run(r, size=10.5, bold=False, color=INK)

        # 例句：英文一段 + 中文一段（复刻原表1 排版）
        p_en = c[3].paragraphs[0]
        r = p_en.add_run(en)
        set_run(r, size=10, bold=False, color=INK)
        p_cn = c[3].add_paragraph()
        p_cn.paragraph_format.space_before = Pt(0)  # 原表：中文例句段前距 0
        r = p_cn.add_run(cn)
        set_run(r, size=9.5, bold=False, color=GREY)

    # 原表：表格与下一个主题标题之间没有空段落，直接相接（更紧凑）

# --- 结尾说明
add_para(doc, f"本补遗共收录 {total_words} 个词目。", size=10.5, bold=True, color=NAVY, before=120, after=40)
add_para(doc,
         "音标为英式（IPA）标注，斜杠内的 (r) 表示英语中可不发音的 r。"
         "完整差集（481 词）见工作区 missing_tight.txt，需要全量版可再扩展。",
         size=10, bold=False, color=GREY, after=40)

out = WS + "/KET词汇补遗_v4.docx"
doc.save(out)
print("saved", out, "total entries", total_words, "topics", len(DATA))
