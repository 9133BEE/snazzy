
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$csv1 = Import-Csv "C:\Users\Authma\Desktop\Claude\受眾包\受眾包all\ExportBlock-6baf9576-a3ac-45b6-80d3-fc9f305c0513-Part-1\TA 233b1307b65f80da8e48e3942cdce925_all.csv" -Encoding UTF8
$csv2 = Import-Csv "C:\Users\Authma\Desktop\Claude\受眾包\受眾包_可&優.csv" -Encoding UTF8
$all = $csv1 + $csv2

$locPat = '公里|車站|捷運|高鐵|People who like|網站訪客|貼文互動|粉專互動|像素|名單|附近|^\+|路[一二三四五六七八九十\d]|街[一二三四五六七八九十\d]|[一二三四五六七八九十\d]+段$|[一二三四五六七八九十\d]+號$|路$|街$|巷$|弄$|[區縣鄉鎮]$|台中|臺中|台北|臺北|新北|桃園|新竹|苗栗|彰化|南投|雲林|嘉義|台南|高雄|基隆|宜蘭|花蓮|台東|澎湖|金門|連江|清水|神岡|潭子|大甲|豐原|東勢|石岡|后里|外埔|梧棲|龍井|大肚|沙鹿|霧峰|烏日|南屯|西屯|北屯|太平|大里|大雅|和平|林口|五股|板橋|三重|中和|永和|新莊|土城|樹林|淡水|新店|蘆洲|汐止|三峽|鶯歌|Taichung|\d+-\d+\s*/'
$locCatPat = '公里|\+\d|地區|^\d'

$zhEn = @{
    "人工智慧"="Artificial intelligence";"機器學習"="Machine learning";"大數據"="Big data"
    "雲端運算"="Cloud computing";"資訊安全"="Cybersecurity";"網路安全"="Cybersecurity"
    "物聯網"="Internet of things";"軟體工程"="Software engineering";"電子工程"="Electrical engineering"
    "資訊科技"="Information technology";"數位轉型"="Digital transformation";"自動化技術"="Automation"
    "電動車"="Electric vehicle";"半導體"="Semiconductor";"自動化"="Automation"
    "家庭自動化"="Smart home automation";"穿戴式科技"="Wearable technology";"智慧家居"="Smart home"
    "行動應用程式開發"="Mobile application development";"開放原始碼"="Open source software"
    "伸展"="Stretching";"科技新聞"="Tech news";"產業科技"="Industrial technology"
    "電子元件"="Electronic components";"科技傳播"="Science communication"
    "投資理財"="Personal finance";"投資"="Investment";"房地產投資"="Real estate investing"
    "投資報酬率"="Return on investment";"指數股票型基金"="Exchange-traded fund"
    "被動收入"="Passive income";"股票市場"="Stock market";"股票"="Stock market"
    "外匯市場"="Foreign exchange market";"加密貨幣"="Cryptocurrency";"比特幣"="Bitcoin"
    "以太幣"="Ethereum";"區塊鏈"="Blockchain";"交易策略"="Trading strategy"
    "高淨值人士"="High-net-worth individual";"財富管理"="Wealth management";"財富"="Wealth"
    "財務規劃"="Financial planning";"個人理財"="Personal finance";"理財"="Personal finance"
    "退休規劃"="Retirement planning";"退休社區"="Retirement community";"退休金"="Pension"
    "金融市場"="Financial market";"金融科技"="Financial technology"
    "演算法交易"="Algorithmic trading";"天使投資"="Angel investing";"創投融資"="Venture capital financing"
    "認證理財規劃師"="Certified Financial Planner";"私人銀行"="Private banking"
    "報稅服務"="Tax preparation";"資產管理"="Asset management";"對沖基金"="Hedge fund"
    "私募股權"="Private equity";"信託"="Trust fund";"股票經紀人"="Stock broker"
    "房地產"="Real estate";"不動產"="Real estate";"房屋"="House";"公寓"="Apartment"
    "公寓大樓"="Condominium";"住宅區"="Residential area";"商用不動產"="Commercial property"
    "房地產開發"="Real estate development";"房屋抵押貸款"="Mortgage loan";"租賃"="Leasing"
    "住家翻新"="Home renovation";"首次購房"="First-time home buyer"
    "房地產開發商"="Real estate developer";"連棟住宅"="Townhouse"
    "商業"="Business";"企業家"="Entrepreneurship";"新創公司"="Startups";"小型企業"="Small business"
    "中小型企業"="Small and medium enterprises";"外包"="Outsourcing"
    "加盟連鎖"="Franchising";"商業模式"="Business model";"企業資源計劃"="Enterprise resource planning"
    "專業服務"="Professional services";"商機"="Business opportunities"
    "企業孵化器"="Business incubator";"創業加速器"="Seed accelerator"
    "共用工作空間"="Co-working";"成人教育"="Adult education";"商業社交"="Business networking"
    "領導力發展"="Leadership development";"創意創業"="Entrepreneurship"
    "員工敬業度"="Employee engagement";"開發潛在顧客"="Lead generation"
    "行銷自動化"="Marketing automation";"銷售管理"="Sales management"
    "企業管理"="Business management";"人力資源管理"="Human resource management"
    "商業決策者"="Business decision maker";"高階管理人員"="Senior management";"會計"="Accounting"
    "業務開發"="Business development";"採購管理"="Procurement";"支付服務"="Payment service"
    "自由工作者市場"="Freelance marketplace";"運輸和物流"="Logistics";"餐旅業"="Hospitality industry"
    "顧問諮詢"="Consulting";"創新"="Innovation";"用戶體驗設計"="User experience design"
    "專案管理"="Project management";"工作父母"="Working parent";"居家辦公"="Home office"
    "銀行業"="Banking";"網路銀行"="Online banking";"行動支付"="Mobile payment";"信用卡"="Credit card"
    "貸款"="Loan";"保險"="Insurance";"人壽保險"="Life insurance";"風險管理"="Risk management"
    "行銷"="Marketing";"數位行銷"="Digital marketing";"社群媒體行銷"="Social media marketing"
    "品牌管理"="Brand management";"聯盟行銷"="Affiliate marketing";"網路廣告"="Online advertising"
    "網紅行銷"="Influencer marketing"
    "汽車"="Automobile";"豪華轎車"="Luxury vehicle";"跑車"="Sports car"
    "摩托車"="Motorcycle";"自行車"="Cycling";"公共運輸"="Public transport";"通勤"="Commuting"
    "摩托車配件"="Motorcycle accessories";"汽車零件"="Car accessories";"汽車維修"="Auto repair"
    "旅遊"="Travel";"自助旅行"="Independent travel";"文化旅遊"="Cultural travel"
    "奢華旅遊"="Luxury travel";"生態旅遊"="Ecotourism";"背包旅行"="Backpacking"
    "文化觀光"="Cultural tourism";"城市探險"="Urban exploration";"公路旅行"="Road trip"
    "周末小旅行"="Weekend getaway";"宅度假"="Staycation";"露營車旅遊"="RV travel"
    "奢華渡假村"="Luxury resort";"國家公園"="National parks";"冒險旅遊"="Adventure travel"
    "機場貴賓室"="Priority Pass";"飛行常客"="Frequent flyer programs"
    "美食"="Food";"烹飪"="Cooking";"咖啡"="Coffee";"葡萄酒"="Wine";"啤酒"="Beer"
    "健康飲食"="Healthy eating";"外食"="Dining out";"美食家"="Gourmet";"品酒"="Wine tasting"
    "精緻料理"="Fine dining";"精品咖啡"="Specialty coffee";"精釀啤酒"="Craft Beer and Brewing"
    "手工啤酒"="Craft Beer and Brewing";"健康食譜"="Healthy recipes";"茶葉"="Tea"
    "英式早餐"="Full English breakfast";"早餐"="Breakfast";"餐飲業"="Food service"
    "美食愛好者"="Gourmet food lovers";"優質食品"="Quality foods";"小型啤酒廠"="Microbrewery"
    "健康零食"="Healthy snacks";"即時食品"="Convenience food";"食材"="Food ingredients"
    "在家下廚"="Home cooking";"農場到餐桌"="Farm-to-table";"有機食品"="Organic food"
    "健身"="Fitness";"瑜珈"="Yoga";"皮拉提斯"="Pilates";"跑步"="Running";"登山"="Hiking"
    "露營"="Camping";"戶外活動"="Outdoor recreation";"羽球"="Badminton";"高爾夫"="Golf"
    "健康照護"="Health care";"醫療"="Medicine";"營養"="Nutrition";"保健食品"="Health food"
    "維他命"="Vitamins";"蛋白質"="Protein";"身心健康"="Mental health";"心理學"="Psychology"
    "教育"="Education";"親子教育"="Parenting";"幼兒教育"="Early childhood education"
    "高等教育"="Higher education";"雙語教育"="Bilingual education";"線上學習"="Online learning"
    "遠距教學"="Distance education";"職業教育"="Vocational education"
    "語言教育"="Language education";"護理教育"="Nursing education"
    "教育資源"="Educational resources";"益智玩具"="Educational toys";"課外活動"="Extracurricular activities"
    "繼續教育"="Continuing education";"專業發展"="Professional development"
    "早教中心"="Early Learning Centre";"教養方式"="Parenting styles";"育兒"="Parenting"
    "室內設計"="Interior design";"居家裝修"="Home improvement";"家居裝潢"="Home decoration"
    "廚房設計"="Kitchen design";"園藝"="Gardening";"大掃除"="Spring cleaning"
    "家務管理"="Home organization";"居家裝飾"="Home decor";"北歐風設計"="Scandinavian design"
    "組合屋"="Prefabricated home";"永續建築"="Sustainable architecture";"燈具"="Light fixture"
    "庭院家具"="Garden furniture";"入厝派對"="Housewarming party";"露台"="Terrace"
    "藝術"="Art";"音樂"="Music";"攝影"="Photography";"設計"="Design";"時尚"="Fashion"
    "當代藝術"="Contemporary art";"藝術展覽"="Art exhibitions";"藝術教育"="Art education"
    "數位藝術"="Digital art";"自然攝影"="Nature photography";"木工"="Woodworking"
    "現場音樂"="Live music";"現場活動"="Live events";"電影愛好者"="Movie lovers"
    "派對策劃"="Party planning";"社交"="Socializing";"藝術拍賣"="Art auction"
    "美容"="Beauty";"保養"="Skin care";"化妝品"="Cosmetics";"手工藝"="Handicraft"
    "女裝品牌"="Women's clothing";"男裝品牌"="Men's clothing";"永續時尚"="Sustainable fashion"
    "男性修容"="Men's grooming";"狗狗美容"="Dog grooming";"狗狗散步"="Dog walking"
    "高級俱樂部"="Country club";"豪奢生活方式"="Luxury lifestyle"
    "精品酒店"="Boutique hotel";"收藏品"="Collectibles";"慈善家"="Philanthropy"
    "收藏"="Collectibles"
    "電影"="Film";"電視"="Television";"遊戲"="Video game";"電競"="Esports"
    "娛樂新聞"="Entertainment news";"名人娛樂新聞"="Celebrity news"
    "桌上遊戲"="Tabletop games";"串流"="Streaming";"網路娛樂"="Online entertainment"
    "爆紅影片"="Viral video";"隨選影片"="Video on demand";"電視製作"="Television production"
    "網路購物"="Online shopping";"電商"="E-commerce";"現金回饋"="Cashback"
    "買一送一"="Buy one get one free";"會員集點"="Loyalty program"
    "寵物"="Pet";"環保"="Sustainability";"慈善"="Philanthropy"
    "永續農業"="Sustainable agriculture";"農夫市集"="Farmers market";"養殖漁業"="Fish farming"
    "農業"="Agriculture";"商業捕魚"="Commercial fishing"
    "製造業"="Manufacturing";"工業工程"="Industrial engineering";"化學工業"="Chemical industry"
    "紡織工業"="Textile industry";"食品產業"="Food industry"
    "生活品質"="Quality of life";"生活風格"="Lifestyle";"流行文化"="Popular culture"
    "社群媒體"="Social media";"上班族"="White-collar worker"
    "電子產品"="Electronics";"消費電子"="Consumer electronics"
    "智慧型手機"="Smartphone";"筆記型電腦"="Laptop";"平板電腦"="Tablet"
    "時事"="Current events";"新聞媒體"="News media"
    "遺傳學"="Genetics";"基因工程"="Genetic engineering";"神經科學"="Neuroscience"
    "蛋白棒"="Protein bar";"蛋白奶昔"="Protein shake"
    "即時成像相機"="Instant camera";"膠原蛋白"="Collagen";"角蛋白"="Keratin"
    "醸造"="Craft brewing";"釀造"="Craft brewing"
    "跨國公司"="Multinational corporation";"國際商務"="International business"
    "薪水及福利"="Compensation and benefits";"工作服"="Workwear"
    "商業領袖"="Business leaders";"辦公室用品"="Office supplies"
    "行銷服務"="Marketing";"行銷服務和組織"="Marketing"
    "一般健康"="Health";"亞洲文化"="Asian American culture"
    "美國疾病與預防控制中心"="Centers for Disease Control and Prevention"
    "地點"="Places";"旅遊景點和活動"="Tourist attractions"
    "影視獎項"="Award shows";"休閒釣魚"="Recreational fishing"
    "極限運動"="Extreme sports";"電競聯盟和競賽"="Esports"
    "球類運動"="Team sports";"專業健身人士"="Fitness professional"
    "科技品牌"="Technology brands";"聯結車"="Semi-truck"
    "房屋廣告和入口網站"="Real estate listing websites"
    "房屋油漆工和裝修師"="House painters and decorators"
}

# Build EN→ZH reverse map from zhEn
$enZh = @{}
$zhEn.GetEnumerator() | ForEach-Object {
    if (-not $enZh.ContainsKey($_.Value)) { $enZh[$_.Value] = $_.Key }
}

# Parse MD file — augments both zhEn (zh→en) and enZh (en→zh)
$mdPath = "C:\Users\Authma\Desktop\Claude\受眾包\受眾包all\ExportBlock-6baf9576-a3ac-45b6-80d3-fc9f305c0513-Part-1\受眾包 226b1307b65f802b903bc758bd50baf1.md"
if (Test-Path $mdPath) {
    Get-Content $mdPath -Encoding UTF8 | ForEach-Object {
        if ($_ -match '^\s*\d+\.\s+([A-Za-z][A-Za-z0-9\s\&\.\-_\/]*?)\s*[\-（(](.+)') {
            $enPart = ($matches[1].Trim() -replace '\*+','')
            $zhRaw  = ($matches[2] -replace '\*+','' -replace '[）)\]]+.*$','')
            $zhParts = $zhRaw -split '[、,，]' | ForEach-Object {
                $_.Trim() -replace '（.*$','' -replace '\(.*$','' -replace '\*+',''
            } | Where-Object { $_.Length -gt 1 }
            foreach ($zp in $zhParts) {
                if (-not $zhEn.ContainsKey($zp)) { $zhEn[$zp] = $enPart }
            }
            if ($zhParts.Count -gt 0 -and -not $enZh.ContainsKey($enPart)) {
                $enZh[$enPart] = $zhParts[0]
            }
        }
    }
}

$catGroup = @{
    "電腦和電子產品"="科技";"消費電子產品"="科技";"電腦硬體"="科技";"資訊技術"="科技"
    "軟體"="科技";"行動裝置"="科技";"計算"="科技";"科學"="科技";"工程"="科技"
    "工業材料和設備"="科技";"產業"="科技";"網站"="科技";"電信"="科技";"智慧型手機"="科技"
    "金屬與礦物"="科技";"能源與公用事業"="科技";"社群媒體"="科技"
    "銀行業"="金融";"投資"="金融";"金融"="金融";"借貸"="金融";"保險"="金融"
    "貨幣"="金融";"信貸和借貸"="金融";"經濟"="金融";"商業和金融業"="金融"
    "商業金融"="金融";"商業和金融"="金融"
    "商業與財務"="商業";"行銷"="商業";"工作機會"="商業";"零售業者"="商業"
    "零售"="商業";"商業活動"="商業";"就業"="商業";"職業訓練"="商業"
    "待售"="商業";"設計"="商業";"製造業"="商業";"郵遞"="商業";"優惠券和折扣"="商業"
    "公司"="商業"
    "美食和餐飲"="美食";"食物和飲品"="美食";"餐廳"="美食";"酒精飲料"="美食"
    "非酒精飲料"="美食";"咖啡"="美食";"葡萄酒"="美食";"食品零售業者"="美食"
    "烘焙食品"="美食";"乳製品"="美食";"飲食"="美食";"餐飲"="美食";"烹飪"="美食"
    "啤酒"="美食";"酒類"="美食"
    "房地產"="房地產興趣";"居家裝潢"="家居";"家居和園藝用品"="家居";"建築"="房地產興趣"
    "建築業"="房地產興趣";"室內裝潢"="家居";"廚房與餐廚"="家居";"住家修繕"="家居"
    "園藝"="家居";"興建和維修"="家居";"家用電器"="家居";"清潔服務"="家居"
    "居家和園藝用品"="家居"
    "服飾"="時尚美妝";"化妝品"="時尚美妝";"美妝"="時尚美妝";"護髮保養"="時尚美妝"
    "飾品配件"="時尚美妝";"手錶"="時尚美妝";"盥洗用品"="時尚美妝"
    "美容服務"="時尚美妝";"身體藝術"="時尚美妝";"鞋"="時尚美妝";"時尚品牌"="時尚美妝"
    "個人護理"="時尚美妝"
    "健身"="健康運動";"體育運動"="健康運動";"運動用品"="健康運動";"水上運動"="健康運動"
    "舉重"="健康運動";"自行車"="健康運動";"滑雪和單板滑雪"="健康運動";"心理學"="健康運動"
    "電玩遊戲"="娛樂";"電影和電視"="娛樂";"音樂"="娛樂";"遊戲"="娛樂"
    "串流服務"="娛樂";"媒體刊物"="娛樂";"電視節目"="娛樂";"電影"="娛樂"
    "流行文化"="娛樂";"娛樂與媒體"="娛樂";"表演藝術"="娛樂";"動畫和漫畫"="娛樂"
    "視覺藝術"="娛樂";"藝術"="娛樂";"傳播"="娛樂";"博奕"="娛樂";"雜誌"="娛樂"
    "新聞服務"="娛樂";"文學"="娛樂";"電視和影片"="娛樂";"攝影"="娛樂"
    "手工藝"="娛樂";"工藝用品"="娛樂";"band"="娛樂"
    "旅遊觀光"="旅遊生活";"旅遊觀光業"="旅遊生活";"戶外活動"="旅遊生活";"休閒"="旅遊生活"
    "活動"="旅遊生活";"露營"="旅遊生活";"航空旅行"="旅遊生活";"住宿"="旅遊生活"
    "生活風格內容"="旅遊生活";"地標"="旅遊生活";"婚禮"="旅遊生活"
    "俱樂部和興趣取向協會"="旅遊生活"
    "教育"="教育";"小學和中學教育"="教育";"高等教育"="教育";"兒童和親子教育"="教育"
    "玩具"="教育"
    "農業"="農業自然";"植物與花卉"="農業自然";"動物"="農業自然";"寵物"="農業自然"
    "寵物用品"="農業自然"
    "社會觀念"="社會";"政治"="社會";"社會公益理念"="社會";"全球"="社會"
    "文化"="社會";"哲學"="社會";"人際關係和身分認同"="社會";"政府機構"="社會"
    "法律和政府"="社會"
    "交通工具"="交通";"汽車"="交通";"汽車零件和服務"="交通";"交通運輸"="交通"
    "行李"="交通";"船舶"="交通"
}

$richPat  = 'Private bank|Wealth management|High.net.worth|Luxury|Fine dining|Art auction|Country club|Boutique hotel|Frequent flyer|Priority Pass|Golf|Philanthropy|Luxury travel|Luxury resort|Luxury vehicle|Collectible|Wine tasting|Gourmet|Sotheby|Christie|Hedge fund|Private equity|Angel invest|Venture capital|Tax prep|Certified Financial|Asset management|Trust fund|Premium|First class|Exclusive'
$youngPat  = 'Netflix|Spotify|TikTok|Instagram|YouTube|Esport|Streaming|K.pop|Pop culture|Anime|Manga|Streetwear|Sneaker|Music festival|Concert|Social media|Online entertain|Video game|Mobile app|Cryptocurrency|Bitcoin|Ethereum|Blockchain|Startup|Digital art|Gaming|Podcast|CrossFit|Plant.based|Vegan|Stand.up|Comedy|Craft [Bb]eer|Adventure travel|NFT'
$rePat     = 'Real estate|Mortgage|Home buyer|Townhouse|Condominium|Residential area|Commercial property|Renovation|Investment club|Open house|Rental listing|Real estate development|Real estate developer'
$techPat   = 'Software development|DevOps|Linux|Open.source|Machine learning|Artificial intelligence|Cloud computing|Cybersecurity|Consumer electronics|TechRadar|Cryptocurrency|Bitcoin|Blockchain|Esport|Gaming|Unity|Software engineering|Electronic component|Industrial technology|Science communication|Physical security|Android|Gadget|Geek|Programming|Big data|Data science|Stack Overflow|GitHub'
$civilPat  = 'Public administration|Retirement planning|Life insurance|National parks|Cultural tourism|Art exhibition|Continuing education|Adult education|Nursing|Public service|Government|Social work|Community|Non.profit|Charity|Volunteer|Historical|Heritage|Pension|Civil service'

function TranslateTag($cp) {
    if ($cp -match '^[A-Za-z0-9][A-Za-z0-9\s&\.\-_\(\)\/]*$') { return $cp }
    foreach ($zk in $zhEn.Keys) { if ($cp -eq $zk) { return $zhEn[$zk] } }
    foreach ($zk in $zhEn.Keys) { if ($cp.StartsWith($zk)) { return $zhEn[$zk] } }
    return $cp
}
function GetZhDesc($enCp, $origCp) {
    if ($enCp -ne $origCp) { return $origCp }
    if ($enZh.ContainsKey($enCp)) { return $enZh[$enCp] }
    return ""
}
function InferCategory([string[]]$tags) {
    $cats = @()
    foreach ($t in $tags) {
        if ($t -imatch 'food|cook|gourmet|dining|restaurant|beer|wine|coffee|tea|cuisine|snack|breakfast|recipe') { $cats += "美食" }
        elseif ($t -imatch 'tech|software|AI|artificial|digital|computer|mobile|app|data|cloud|cyber|IT|machine') { $cats += "科技" }
        elseif ($t -imatch 'invest|financ|bank|stock|crypto|trading|wealth|insurance|fund|equity|currency|mortgage') { $cats += "金融" }
        elseif ($t -imatch 'parent|child|education|learning|school|kids|baby|bilingual|nursery') { $cats += "育兒教育" }
        elseif ($t -imatch 'home|interior|renovation|real estate|property|furniture|garden|decor|house|condo') { $cats += "居家" }
        elseif ($t -imatch 'health|fitness|yoga|running|sport|gym|pilates|nutrition|wellness|medicine') { $cats += "健康運動" }
        elseif ($t -imatch 'fashion|beauty|cosmetic|clothing|style|makeup|skincare|hair') { $cats += "時尚美妝" }
        elseif ($t -imatch 'travel|tourism|trip|vacation|adventure|hotel|resort|flight|backpack') { $cats += "旅遊" }
        elseif ($t -imatch 'business|marketing|entrepreneur|startup|management|sales|brand|retail') { $cats += "商業" }
        elseif ($t -imatch 'gaming|game|esport|movie|music|art|entertainment|streaming|video|anime|concert') { $cats += "娛樂" }
    }
    $uniq = $cats | Select-Object -Unique | Select-Object -First 4
    if ($uniq.Count -gt 0) { return "AI識別：" + ($uniq -join "/") }
    return ""
}

$tagSet  = New-Object 'System.Collections.Generic.HashSet[string]'
$tagList = New-Object 'System.Collections.Generic.List[object]'
$youList = New-Object 'System.Collections.Generic.List[object]'

foreach ($row in $all) {
    $text = $row."興趣"
    if ([string]::IsNullOrWhiteSpace($text)) { continue }

    $iSection = ""
    if ($text -match "興趣[-－]") {
        ($text -split "[\r\n]+") | Where-Object { $_ -match "興趣[-－]" } | ForEach-Object {
            if ($_ -match "興趣[-－](.+)") { $iSection += $matches[1] + "、" }
        }
    } elseif ($text -notmatch "地區[-－]") {
        $iSection = $text
    }
    $iSection = ($iSection -split "•")[0] -replace "任職公司.+$","" -replace "行業類別.+$","" -replace "家長：.+$",""

    $locRaw = ""
    if ($text -match "地區[-－]") {
        ($text -split "[\r\n]+") | Where-Object { $_ -match "地區[-－]" } | Select-Object -First 1 | ForEach-Object {
            if ($_ -match "地區[-－](.+)") { $locRaw = $matches[1] }
        }
    }
    $locClean = ($locRaw -replace '[（(][^）)]*[）)]','') -replace '[,，\s]+','、' -replace '^、+|、+$','' -replace '、{2,}','、'

    $tokens = ($iSection -split "\s*或\s*") | ForEach-Object { $_ -split "[、,，]" } | ForEach-Object {
        # Strip behavior/job-title prefix, then trim
        ($_.Trim() -replace "^(行為|工作職稱|任職公司|學歷|雇主)[：:]\s*","") -replace "^[\+\-\*•\s]+","" -replace "^\d+\.\s*",""
    } | Where-Object { $_.Length -gt 1 -and $_.Length -lt 60 -and $_ -notmatch $locPat }

    foreach ($tok in $tokens) {
        # Strip balanced brackets AND unclosed brackets from copy text
        $cp = ($tok -replace '[（(][^）)]+[）)]','') -replace '[（(][^）)]*$',''
        $cp = $cp.Trim() -replace '\s+',' '
        $cp = $cp -replace '\s+[01]\d[0-3]\d.*$',''   # strip MMDD date+trailing
        $cp = $cp.Trim()
        if ($cp.Length -lt 2 -or $cp -match $locPat -or $cp -match '^\d+$') { continue }
        if ($cp -match '也必須符合|^\d+%|^緯度\s|行業類別：|生活要事：|職稱：') { continue }

        $mc = "其他"
        if ($tok -match '[（(]([^）)]+)[）)]') {
            $rawCat = $matches[1].Trim()
            if ($rawCat -notmatch $locCatPat) {
                $mc = if ($catGroup.ContainsKey($rawCat)) { $catGroup[$rawCat] } else { $rawCat }
            }
        }
        $enCp   = TranslateTag $cp
        $zhDesc = GetZhDesc $enCp $cp
        $sp = ""
        if      ($enCp -imatch $richPat)  { $sp="rich"   }
        elseif  ($enCp -imatch $youngPat) { $sp="young"  }
        elseif  ($enCp -imatch $rePat)    { $sp="re"     }
        elseif  ($enCp -imatch $techPat)  { $sp="tech"   }
        elseif  ($enCp -imatch $civilPat) { $sp="civil"  }

        $key = $enCp + "|" + $mc
        if ($tagSet.Add($key)) {
            $tagList.Add([PSCustomObject]@{cp=$enCp; zh=$zhDesc; cat=$mc; sp=$sp})
        }
    }

    if ($row."成效" -match "優" -and $iSection.Length -gt 3) {
        $taRaw  = [string]$row."TA"
        $taFmt  = ($taRaw -split "[\r\n]+" | Where-Object { $_.Trim().Length -gt 0 -and $_.Trim() -notmatch '^\d+$' }) -join " / "
        $pkgTagSet = New-Object 'System.Collections.Generic.HashSet[string]'
        $pkgTags   = New-Object 'System.Collections.Generic.List[string]'

        $ptokens = ($iSection -split "\s*或\s*") | ForEach-Object { $_ -split "[、,，]" } | ForEach-Object {
            ($_.Trim() -replace "^(行為|工作職稱|任職公司|學歷|雇主)[：:]\s*","") -replace "^[\+\-\*•\s]+","" -replace "^\d+\.\s*",""
        } | Where-Object { $_.Length -gt 1 -and $_.Length -lt 60 -and $_ -notmatch $locPat }

        foreach ($ptok in $ptokens) {
            $pcp = ($ptok -replace '[（(][^）)]+[）)]','') -replace '[（(][^）)]*$',''; $pcp = $pcp.Trim()
            $pcp = ($pcp -replace '\s+[01]\d[0-3]\d.*$','').Trim()
            if ($pcp.Length -lt 2 -or $pcp -match $locPat -or $pcp -match '^\d+$') { continue }
            if ($pcp -match '也必須符合|^\d+%|^緯度\s|行業類別：|生活要事：|職稱：') { continue }
            $pEn = TranslateTag $pcp
            $pmc = "其他"
            if ($ptok -match '[（(]([^）)]+)[）)]') {
                $pRaw = $matches[1].Trim()
                if ($pRaw -notmatch $locCatPat) { $pmc = if ($catGroup.ContainsKey($pRaw)) { $catGroup[$pRaw] } else { $pRaw } }
            }
            # Store as "tag::category"
            if ($pkgTagSet.Add($pEn)) { $pkgTags.Add("${pEn}::${pmc}") }
        }

        $projRaw = ([string]$row."專案").Trim()
        if ([string]::IsNullOrWhiteSpace($projRaw)) {
            $plainTags = $pkgTags | ForEach-Object { ($_ -split "::")[0] }
            $autoName = InferCategory ($plainTags | Select-Object -First 8)
            $projRaw = if ($autoName) { $autoName } else { "未命名" }
        }

        $youList.Add([PSCustomObject]@{
            ta   = $taFmt
            proj = $projRaw
            tags = $pkgTags -join "||"
            loc  = $locClean
            lx   = ([string]$row."類型").Trim()
        })
    }
}

# Manually curated tags not in CSV data
$manualTags = @(
    @{cp="Business leaders";       zh="商業領袖";              cat="商業";         sp="hidden"}
    @{cp="Office supplies";        zh="辦公室用品";            cat="商業";         sp="hidden"}
    @{cp="Marketing";              zh="行銷服務和組織";        cat="商業";         sp="hidden"}
    @{cp="Health";                 zh="一般健康";              cat="健康運動";     sp="hidden"}
    @{cp="Asian American culture"; zh="亞洲文化";              cat="社會";         sp="hidden"}
    @{cp="Centers for Disease Control and Prevention"; zh="美國疾病與預防控制中心"; cat="健康運動"; sp="hidden"}
    @{cp="Places";                 zh="地點";                  cat="旅遊生活";     sp="hidden"}
    @{cp="Tourist attractions";    zh="旅遊景點和活動";        cat="旅遊生活";     sp="hidden"}
    @{cp="Award shows";            zh="影視獎項";              cat="娛樂";         sp="young"}
    @{cp="Recreational fishing";   zh="休閒釣魚";              cat="旅遊生活";     sp="hidden"}
    @{cp="Extreme sports";         zh="極限運動";              cat="健康運動";     sp="young"}
    @{cp="Esports leagues and competitions"; zh="電競聯盟和競賽"; cat="娛樂";      sp="young"}
    @{cp="Team sports";            zh="球類運動";              cat="健康運動";     sp="hidden"}
    @{cp="Fitness professional";   zh="專業健身人士";          cat="健康運動";     sp="hidden"}
    @{cp="Technology brands";      zh="科技品牌";              cat="科技";         sp="tech"}
    @{cp="Semi-truck";             zh="聯結車";                cat="交通";         sp="hidden"}
    @{cp="Real estate listing websites"; zh="房屋廣告和入口網站"; cat="房地產興趣"; sp="re"}
    @{cp="House painters and decorators"; zh="房屋油漆工和裝修師"; cat="房地產興趣"; sp="re"}
    @{cp="Real estate web";        zh="房地產行銷和培訓";      cat="房地產興趣";   sp="re"}
    @{cp="Viral video";            zh="爆紅影片";              cat="娛樂";         sp="young"}
    # --- 家庭、健康、教育 ---
    @{cp="Household";                  zh="居家用品零售商、家庭";        cat="家居";         sp="hidden"}
    @{cp="LifeStyle Home";             zh="生活風格家居";                cat="家居";         sp="hidden"}
    @{cp="Motherhood";                 zh="育兒雜誌";                    cat="教育";         sp="hidden"}
    @{cp="Parenting";                  zh="嬰兒副食品";                  cat="教育";         sp="hidden"}
    @{cp="Family travel";              zh="家庭桌遊";                    cat="旅遊生活";     sp="hidden"}
    @{cp="Kids activities";            zh="兒童活動";                    cat="教育";         sp="hidden"}
    @{cp="Cultural";                   zh="交換學生";                    cat="教育";         sp="hidden"}
    @{cp="Kitchen appliances";         zh="廚房家電、家電品牌和零售商";  cat="家居";         sp="hidden"}
    @{cp="Design studios";             zh="嬰幼兒和孕婦用品店";          cat="教育";         sp="hidden"}
    @{cp="Healthy";                    zh="健康的生活習慣";              cat="健康運動";     sp="hidden"}
    @{cp="Nutritional supplements";    zh="維他命和營養補充品";          cat="健康運動";     sp="hidden"}
    @{cp="Nursing";                    zh="護理教育";                    cat="教育";         sp="civil"}
    @{cp="BACK";                       zh="健康食品品牌";                cat="健康運動";     sp="hidden"}
    @{cp="Online learning";            zh="教育技術學";                  cat="教育";         sp="hidden"}
    @{cp="Nutrition";                  zh="人體營養";                    cat="健康運動";     sp="hidden"}
    @{cp="House work";                 zh="做家事";                      cat="家居";         sp="hidden"}
    @{cp="HVAC";                       zh="暖通空調";                    cat="家居";         sp="hidden"}
    @{cp="Protein";                    zh="蛋白棒、蛋白奶昔、蛋白質";   cat="健康運動";     sp="hidden"}
    @{cp="Cabinetry";                  zh="櫥櫃";                        cat="家居";         sp="hidden"}
    @{cp="Healthy recipes";            zh="健康飲食的食譜";              cat="美食";         sp="hidden"}
    @{cp="Garden furniture";           zh="庭院家具";                    cat="家居";         sp="hidden"}
    @{cp="Gazebo";                     zh="花園建築和裝飾";              cat="家居";         sp="hidden"}
    # --- 投資、工作 ---
    @{cp="Payment service provider";   zh="支付服務商";                  cat="金融";         sp="hidden"}
    @{cp="Project management software";zh="專案管理軟體";                cat="商業";         sp="tech"}
    @{cp="Electrical wiring";          zh="電氣配線";                    cat="家居";         sp="hidden"}
    @{cp="Trade shows";                zh="設計貿易展和專業組織";        cat="商業";         sp="hidden"}
    @{cp="Employee engagement";        zh="員工敬業度";                  cat="商業";         sp="hidden"}
    @{cp="Medical Sales Representative";zh="醫療業務代表";               cat="商業";         sp="hidden"}
    @{cp="Freight transport";          zh="運輸和物流服務和組織";        cat="商業";         sp="hidden"}
    @{cp="Work from home";             zh="遠距工作";                    cat="商業";         sp="tech"}
    @{cp="Customer support";           zh="客戶支援";                    cat="商業";         sp="hidden"}
    # --- 高收高消費 ---
    @{cp="Christie's";                 zh="佳士得、藝術奢侈品拍賣";     cat="旅遊生活";     sp="rich"}
    @{cp="Sotheby's";                  zh="蘇富比、藝術奢侈品拍賣";     cat="旅遊生活";     sp="rich"}
    @{cp="Fine dining";                zh="高級餐飲愛好者";              cat="美食";         sp="rich"}
    @{cp="Boutique";                   zh="女性精品服飾零售商";          cat="時尚美妝";     sp="hidden"}
    @{cp="Luxury";                     zh="精品生活風格";                cat="旅遊生活";     sp="rich"}
    @{cp="Jewelry";                    zh="客製化禮品或服裝";            cat="時尚美妝";     sp="hidden"}
    # --- 美食、酒、餐飲 ---
    @{cp="FineDiningLovers";           zh="精緻料理愛好者";              cat="美食";         sp="rich"}
    @{cp="Foodie";                     zh="美食家";                      cat="美食";         sp="hidden"}
    @{cp="Healthy cooking";            zh="食品媒體和內容";              cat="美食";         sp="hidden"}
    @{cp="Local events";               zh="葡萄酒相關內容";              cat="美食";         sp="hidden"}
    @{cp="KIND Healthy Snacks";        zh="健康零食品牌";                cat="美食";         sp="hidden"}
    @{cp="Food trucks";                zh="飲食業";                      cat="美食";         sp="hidden"}
    @{cp="Tea leaf";                   zh="茶葉品牌";                    cat="美食";         sp="hidden"}
    @{cp="Control system";             zh="酒類商品零售商";              cat="美食";         sp="hidden"}
    @{cp="Foodservice";                zh="餐飲業";                      cat="美食";         sp="hidden"}
    @{cp="LifeStyle Food";             zh="輕食、均衡飲食";              cat="美食";         sp="hidden"}
    # --- 購物、網路、美容 ---
    @{cp="Weekend";                    zh="男裝品牌、女裝品牌";          cat="時尚美妝";     sp="hidden"}
    @{cp="Collectible";                zh="收藏用玩具、集換卡片遊戲";    cat="娛樂";         sp="hidden"}
    @{cp="Fragrance oil";              zh="香精油";                      cat="時尚美妝";     sp="hidden"}
    @{cp="HydraFacial";                zh="海菲秀";                      cat="時尚美妝";     sp="hidden"}
    @{cp="Fashion Tech";               zh="時尚機構";                    cat="時尚美妝";     sp="hidden"}
    @{cp="Video game culture";         zh="電子遊戲文化";                cat="娛樂";         sp="young"}
    @{cp="Digital media";              zh="數位媒體";                    cat="娛樂";         sp="hidden"}
    @{cp="Work from Home";             zh="遠距工作";                    cat="商業";         sp="tech"}
    @{cp="DeFi";                       zh="時裝周、時裝秀";              cat="時尚美妝";     sp="young"}
    # --- 生活、旅遊、藝文 ---
    @{cp="Lifestyle brand";            zh="生活風格品牌";                cat="旅遊生活";     sp="hidden"}
    @{cp="Lifestyle Travel";           zh="貼近日常、深入在地";          cat="旅遊生活";     sp="hidden"}
    @{cp="Home and garden";            zh="生活風格和家務內容";          cat="家居";         sp="hidden"}
    @{cp="Lifestyle content";          zh="節儉生活風格";                cat="旅遊生活";     sp="hidden"}
    @{cp="Outdoor living";             zh="戶外生活";                    cat="旅遊生活";     sp="hidden"}
    @{cp="Hiking equipment";           zh="健行和露營裝備";              cat="旅遊生活";     sp="hidden"}
    @{cp="Travel Deals";               zh="旅行優惠";                    cat="旅遊生活";     sp="hidden"}
    @{cp="Brunch";                     zh="派對策劃";                    cat="旅遊生活";     sp="hidden"}
    @{cp="Happy hour";                 zh="歡樂時光";                    cat="旅遊生活";     sp="hidden"}
    @{cp="Christmas Holiday";          zh="聖誕假期";                    cat="旅遊生活";     sp="hidden"}
    @{cp="Trails";                     zh="健行和露營、登山步道";        cat="旅遊生活";     sp="hidden"}
    @{cp="Birdwatching";               zh="觀鳥";                        cat="旅遊生活";     sp="hidden"}
    @{cp="Tiny House Movement";        zh="居家風格和生活類型";          cat="家居";         sp="hidden"}
    @{cp="Skyscanner";                 zh="機票比價平台";                cat="旅遊生活";     sp="hidden"}
    @{cp="Cheapflights";               zh="機票比價平台";                cat="旅遊生活";     sp="hidden"}
    @{cp="Self-guided tour";           zh="自助導覽";                    cat="旅遊生活";     sp="hidden"}
    @{cp="Chinese Culture";            zh="中華文化";                    cat="社會";         sp="civil"}
    @{cp="Art gallery";                zh="國家藝廊";                    cat="娛樂";         sp="civil"}
    @{cp="High-end audio";             zh="高端音響";                    cat="科技";         sp="rich"}
    @{cp="Weekend Trips";              zh="週末旅遊";                    cat="旅遊生活";     sp="hidden"}
    # --- 房地產興趣 ---
    @{cp="Real estate";                zh="住宅房地產、住宅房地產經紀商"; cat="房地產興趣";  sp="re"}
    @{cp="Investment club";            zh="房地產投資俱樂部";            cat="房地產興趣";   sp="re"}
    @{cp="Open house";                 zh="建築媒體";                    cat="房地產興趣";   sp="re"}
    @{cp="Rental listings";            zh="待租房屋資訊和內容";          cat="房地產興趣";   sp="re"}
    @{cp="Benefit";                    zh="退伍軍人福利";                cat="社會";         sp="civil"}
    @{cp="Public administration";      zh="市政事務";                    cat="社會";         sp="civil"}
    # --- 企業、商務、投資 ---
    @{cp="Retirement Insurance Benefits";zh="國民年金、退休保險";        cat="金融";         sp="civil"}
    @{cp="Lending";                    zh="信貸服務和組織";              cat="金融";         sp="hidden"}
    @{cp="Foreign company";            zh="通貨";                        cat="金融";         sp="hidden"}
    @{cp="MBA";                        zh="行動銀行";                    cat="金融";         sp="civil"}
    # --- 科技業興趣 ---
    @{cp="TechRadar";                  zh="全球科技媒體網站";            cat="科技";         sp="tech"}
    @{cp="Geek.com";                   zh="美國科技新聞網站";            cat="科技";         sp="tech"}
    @{cp="Physical security";          zh="保護人員和財產的安全措施";    cat="科技";         sp="tech"}
    @{cp="Top Gear";                   zh="極速誌";                      cat="交通";         sp="hidden"}
    @{cp="Android Wear";               zh="Android智慧手錶作業系統";    cat="科技";         sp="tech"}
    @{cp="Gadget & Gear";              zh="電子產品、戶外裝備和配件";    cat="科技";         sp="tech"}
    @{cp="KitchenAid";                 zh="家用電器";                    cat="家居";         sp="hidden"}
    @{cp="Consumer Electronics Show";  zh="消費電子展";                  cat="科技";         sp="tech"}
    @{cp="Software development";       zh="基於組件的軟體工程";          cat="科技";         sp="tech"}
    @{cp="GameFAQs";                   zh="電玩遊戲網路媒體";            cat="娛樂";         sp="young"}
    @{cp="Forklift";                   zh="叉車";                        cat="交通";         sp="hidden"}
    @{cp="Family car";                 zh="家用車";                      cat="交通";         sp="hidden"}
    @{cp="Scientific";                 zh="科學媒體和內容";              cat="科技";         sp="tech"}
    @{cp="Enterprise Rent-A-Car";      zh="企業租車";                    cat="交通";         sp="hidden"}
    @{cp="Linux";                      zh="自由和開放原始碼的類Unix作業系統"; cat="科技";    sp="tech"}
    @{cp="Pilates";                    zh="皮拉提斯、健身器材和服裝";    cat="健康運動";     sp="hidden"}
    @{cp="Unity";                      zh="遊戲引擎";                    cat="科技";         sp="tech"}
    @{cp="DevOps";                     zh="開發微運整合";                cat="科技";         sp="tech"}
    # --- 傳產、農林漁牧 ---
    @{cp="Rural";                      zh="農業新聞和內容、農家樂";      cat="農業自然";     sp="hidden"}
    @{cp="Longlining";                 zh="延繩捕魚";                    cat="農業自然";     sp="hidden"}
)
foreach ($mt in $manualTags) {
    $key = $mt.cp + "|" + $mt.cat
    if ($tagSet.Add($key)) {
        $tagList.Add([PSCustomObject]@{cp=$mt.cp; zh=$mt.zh; cat=$mt.cat; sp=$mt.sp})
    }
}

$masterCats = @('科技','金融','商業','美食','房地產興趣','家居','時尚美妝','健康運動','娛樂','旅遊生活','教育','農業自然','社會','交通','其他')
foreach ($tag in $tagList) {
    if ($masterCats -notcontains $tag.cat) {
        $c = $tag.cat
        if      ($c -match '科技|電腦|軟體|資訊|網路|人工智慧') { $tag.cat='科技' }
        elseif  ($c -match '金融|投資|銀行|保險|財務|股票|加密') { $tag.cat='金融' }
        elseif  ($c -match '商業|行銷|企業|零售|就業|管理')     { $tag.cat='商業' }
        elseif  ($c -match '食|飲|餐|咖啡|酒|烹')               { $tag.cat='美食' }
        elseif  ($c -match '房地產|建築|物業')                   { $tag.cat='房地產興趣' }
        elseif  ($c -match '家居|裝潢|園藝|廚房|住家')          { $tag.cat='家居' }
        elseif  ($c -match '時尚|美容|化妝|服飾|美妝')          { $tag.cat='時尚美妝' }
        elseif  ($c -match '健康|運動|體育|健身|醫療|養生')     { $tag.cat='健康運動' }
        elseif  ($c -match '娛樂|電影|音樂|遊戲|藝術|媒體|攝影'){ $tag.cat='娛樂' }
        elseif  ($c -match '旅遊|生活|戶外|活動|露營|航空')     { $tag.cat='旅遊生活' }
        elseif  ($c -match '教育|學習|學校|兒童|親子')          { $tag.cat='教育' }
        elseif  ($c -match '農業|自然|動物|植物|寵物|漁業')     { $tag.cat='農業自然' }
        elseif  ($c -match '社會|文化|政治|公益|宗教')          { $tag.cat='社會' }
        elseif  ($c -match '交通|汽車|運輸|航')                 { $tag.cat='交通' }
        else                                                      { $tag.cat='其他' }
    }
}

Write-Host "Tags: $($tagList.Count)  YOU: $($youList.Count)"

function JsStr($s) { $s -replace '\\','\\' -replace '"','\"' -replace "`r","" -replace "`n"," " }

$dLines = $tagList | ForEach-Object { "{cp:`"$(JsStr $_.cp)`",z:`"$(JsStr $_.zh)`",c:`"$(JsStr $_.cat)`",sp:`"$($_.sp)`"}" }
$jsD = "const D=[" + ($dLines -join ",") + "];"

$yLines = $youList | ForEach-Object { "{ta:`"$(JsStr $_.ta)`",proj:`"$(JsStr $_.proj)`",tags:`"$(JsStr $_.tags)`",loc:`"$(JsStr $_.loc)`",lx:`"$(JsStr $_.lx)`"}" }
$jsY = "const YOU=[" + ($yLines -join ",") + "];"

$catC = @{}
foreach ($r in $tagList) { $c=$r.cat; if(-not $catC[$c]){$catC[$c]=0}; $catC[$c]++ }
$catOpts = ($catC.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    $ck = $_.Key -replace '"','&quot;'; "<option value=`"$ck`">$ck ($($_.Value))</option>"
}) -join "`n"

$lxC = @{}
foreach ($y in $youList) { $l=$y.lx; if($l.Length -gt 0){ if(-not $lxC[$l]){$lxC[$l]=0}; $lxC[$l]++ } }
$lxOpts = ($lxC.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    "<option value=`"$($_.Key)`">$($_.Key) ($($_.Value))</option>"
}) -join "`n"

$nRich   = ($tagList | Where-Object { $_.sp -eq "rich"   }).Count
$nYoung  = ($tagList | Where-Object { $_.sp -eq "young"  }).Count
$nHidden = ($tagList | Where-Object { $_.sp -eq "hidden" }).Count
$nRe     = ($tagList | Where-Object { $_.sp -eq "re"     }).Count
$nTech   = ($tagList | Where-Object { $_.sp -eq "tech"   }).Count
$nCivil  = ($tagList | Where-Object { $_.sp -eq "civil"  }).Count
$n = $tagList.Count; $nc = $catC.Count; $ny = $youList.Count

$htmlPart1 = @"
<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><title>META 受眾工具</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#fcfcfb;--s:#ffffff;--s2:#f4f3f1;--s3:#eff0f2;
  --bd:#e4e1dd;--bd2:#ccc8c3;
  --tx:#2e2c29;--sub:#7d7870;
  --a:#607d96;--al:#eaf0f6;--ab:#a8bed0;--a-dark:#4d6a82;
  --good:#3d6b4f;--goodl:#d8ead2;--goodb:#96c4a8;
  --sp:#5270a0;--spl:#e8edf8;--spb:#9db0d8;
  --rich:#7a5c20;--richl:#fdf3dc;--richb:#d4b06a;
  --hidden:#8a3f3f;--hiddenl:#faeaea;--hiddenb:#c89090;
  --young:#2d6e7e;--youngl:#ddf2f7;--youngb:#7ac0d0;
  --re:#4a6e3a;--rel:#e4f2dc;--reb:#90c080;
  --tech:#484878;--techl:#ecedf8;--techb:#9090c8;
  --radius-card:10px;--radius-chip:6px;--radius-pill:20px;
  --shadow-xs:0 1px 2px rgba(0,0,0,.04);
  --shadow-sm:0 2px 6px rgba(0,0,0,.07),0 1px 2px rgba(0,0,0,.04);
  --shadow-md:0 4px 16px rgba(0,0,0,.10),0 2px 4px rgba(0,0,0,.05);
  --shadow-bar:0 2px 12px rgba(0,0,0,.09),0 1px 3px rgba(0,0,0,.05);
  --transition:all .15s ease;
}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang TC',sans-serif;background:var(--bg);color:var(--tx);font-weight:500;font-size:14px}

/* ── Header ── */
header{background:var(--s);border-bottom:1px solid var(--bd);padding:12px 24px;display:flex;align-items:center;gap:14px;box-shadow:var(--shadow-xs)}
.logo{width:34px;height:34px;background:linear-gradient(135deg,var(--a),var(--a-dark));border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:900;color:#fff;flex-shrink:0;box-shadow:0 2px 6px rgba(96,125,150,.35)}
.htxt h1{font-size:15px;font-weight:700;color:var(--tx);letter-spacing:-.2px}.htxt p{font-size:11px;color:var(--sub);margin-top:2px}
.hstats{margin-left:auto;display:flex;gap:20px;align-items:center}
.hstat{text-align:right}
.hn{font-size:22px;font-weight:800;color:var(--a);line-height:1;letter-spacing:-.5px}
.hl{font-size:10px;color:var(--sub);margin-top:2px;letter-spacing:.2px;text-transform:uppercase;font-weight:600}

/* ── Tabs ── */
.tabs{background:var(--s);border-bottom:1px solid var(--bd);display:flex;padding:0 24px;gap:2px}
.tab{padding:11px 18px;font-size:13px;font-weight:500;color:var(--sub);cursor:pointer;border-bottom:3px solid transparent;margin-bottom:-1px;transition:var(--transition);position:relative}
.tab:hover{color:var(--tx);background:var(--s2);border-radius:6px 6px 0 0}
.tab.on{color:var(--a);border-bottom-color:var(--a);font-weight:700;background:linear-gradient(to bottom,var(--al) 0%,transparent 100%)}

/* ── Sticky toolbar ── */
.bar{background:rgba(255,255,255,.97);padding:10px 20px;border-bottom:1px solid var(--bd);display:flex;gap:8px;flex-wrap:wrap;align-items:center;position:sticky;top:0;z-index:99;box-shadow:var(--shadow-bar);backdrop-filter:blur(6px)}
.sw{position:relative;flex:1;min-width:160px}
.sw input{width:100%;padding:8px 14px 8px 34px;background:var(--s2);border:1.5px solid var(--bd);border-radius:var(--radius-pill);font-size:13px;color:var(--tx);outline:none;transition:var(--transition);font-weight:500}
.sw input::placeholder{color:var(--sub)}
.sw input:focus{border-color:var(--a);background:#fff;box-shadow:0 0 0 3px rgba(96,125,150,.12)}
.sw::before{content:"⌕";position:absolute;left:11px;top:50%;transform:translateY(-50%);font-size:16px;color:var(--sub);pointer-events:none}
select{padding:8px 12px;background:var(--s2);border:1.5px solid var(--bd);border-radius:var(--radius-pill);font-size:12px;color:var(--tx);cursor:pointer;outline:none;font-weight:500;transition:var(--transition)}
select:focus{border-color:var(--a);box-shadow:0 0 0 3px rgba(96,125,150,.12)}
.cnt{font-size:11px;color:var(--sub);white-space:nowrap;margin-left:auto;background:var(--s2);padding:5px 12px;border-radius:var(--radius-pill);border:1px solid var(--bd);font-weight:600}
.ttu{background:none;border:1.5px solid var(--bd);color:var(--sub);width:30px;height:30px;border-radius:50%;font-size:13px;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:var(--transition);line-height:1}
.ttu:hover{background:var(--al);border-color:var(--ab);color:var(--a)}

/* ── Layout ── */
.main{padding:20px;max-width:1280px;margin:0 auto}.pg{display:none}.pg.on{display:block}
.sec{margin-bottom:24px}
.sh{display:flex;align-items:center;gap:10px;margin-bottom:10px}
.st{font-size:11px;font-weight:700;color:var(--a);background:var(--al);border:1.5px solid var(--ab);padding:3px 12px;border-radius:var(--radius-pill);letter-spacing:.4px;text-transform:uppercase}
.sc{font-size:11px;color:var(--sub);font-weight:600}
.ln{flex:1;height:1px;background:linear-gradient(to right,var(--bd),transparent)}

/* ── Tag grid ── */
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:7px}
.card{background:var(--s);border:1px solid var(--bd);border-radius:var(--radius-card);padding:10px 13px;cursor:pointer;transition:var(--transition);position:relative;display:flex;flex-direction:column;gap:3px;box-shadow:var(--shadow-xs)}
.card:hover{border-color:var(--ab);box-shadow:var(--shadow-sm);transform:translateY(-2px)}
.card:active{transform:scale(.97);box-shadow:var(--shadow-xs)}
.te{font-size:12.5px;font-weight:600;color:var(--tx);padding-right:40px;word-break:break-word;line-height:1.4}
.tz{font-size:11px;color:var(--sub);line-height:1.3;font-weight:500}

/* ── Category badges on cards ── */
.tc{font-size:10px;font-weight:700;padding:2px 8px;border-radius:var(--radius-chip);align-self:flex-start;margin-top:3px;letter-spacing:.2px;border:1px solid}
.tc{color:var(--a);background:var(--al);border-color:var(--ab)}
.tc-rich{color:var(--rich)!important;background:var(--richl)!important;border-color:var(--richb)!important}
.tc-young{color:var(--young)!important;background:var(--youngl)!important;border-color:var(--youngb)!important}
.tc-re{color:var(--re)!important;background:var(--rel)!important;border-color:var(--reb)!important}
.tc-tech{color:var(--tech)!important;background:var(--techl)!important;border-color:var(--techb)!important}
.tc-civil{color:var(--sp)!important;background:var(--spl)!important;border-color:var(--spb)!important}
.tc-hidden{color:var(--hidden)!important;background:var(--hiddenl)!important;border-color:var(--hiddenb)!important}
.cb{position:absolute;top:8px;right:8px;background:var(--a);border:none;color:#fff;font-size:10px;padding:3px 8px;border-radius:var(--radius-chip);opacity:0;transition:opacity .15s;cursor:pointer;white-space:nowrap;font-weight:600;box-shadow:var(--shadow-xs)}
.card:hover .cb{opacity:1}

/* ── Package grid (優等受眾包 / 類型分析) ── */
.pgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(330px,1fr));gap:12px}
.pkg{background:var(--s);border:1px solid var(--bd);border-radius:var(--radius-card);overflow:hidden;box-shadow:var(--shadow-xs);transition:var(--transition)}
.pkg:hover{border-color:var(--ab);box-shadow:var(--shadow-sm);transform:translateY(-2px)}
.pkg-header{background:linear-gradient(135deg,var(--s2) 0%,var(--s) 100%);border-bottom:1px solid var(--bd);padding:12px 14px}
.phd{display:flex;align-items:flex-start;gap:8px}
.pbadge{background:var(--goodl);color:var(--good);font-size:10px;font-weight:800;padding:3px 8px;border-radius:var(--radius-chip);border:1px solid var(--goodb);flex-shrink:0;letter-spacing:.3px}
.ptitle{font-size:14px;font-weight:700;color:var(--tx);line-height:1.3}
.psub{font-size:11px;color:var(--sub);margin-top:4px;font-weight:500;line-height:1.5}
.ploc{font-size:10px;color:var(--sub);margin-top:4px;font-weight:500}
.pkg-body{padding:12px 14px}
/* fallback padding for pkgHtml() which nests .phd/.ptags directly inside .pkg */
.pkg>.phd{padding:12px 14px;background:linear-gradient(135deg,var(--s2) 0%,var(--s) 100%);border-bottom:1px solid var(--bd)}
.pkg>.ptags{padding:10px 14px}
/* genPkg() uses .gr-pkg-title + locHtml + desc div + .ptags directly inside .pkg */
.pkg>.gr-pkg-title{padding:12px 14px 0;background:linear-gradient(135deg,var(--s2) 0%,var(--s) 100%);border-bottom:1px solid var(--bd);padding-bottom:10px}
/* padding for inline div children (desc, locHtml) */
.pkg>div:not(.phd):not(.ptags):not(.gr-pkg-title):not(.pkg-header):not(.pkg-body){padding:0 14px}
.ptags{display:flex;flex-wrap:wrap;gap:5px;margin-top:0}
.ptag{font-size:11px;color:var(--tx);background:var(--s2);border:1px solid var(--bd);padding:4px 9px;border-radius:var(--radius-pill);cursor:pointer;transition:var(--transition);font-weight:500;display:inline-flex;align-items:center;gap:4px}
.ptag:hover{background:var(--al);border-color:var(--ab);color:var(--a);transform:translateY(-1px);box-shadow:var(--shadow-xs)}
.ptag-c{font-size:9px;color:var(--sub);background:var(--bg);border-radius:3px;padding:1px 4px;font-weight:400;opacity:.8}
.pkg-perf{display:flex;flex-wrap:wrap;align-items:center;gap:6px;padding:7px 14px 10px;background:var(--al);border-top:1px solid var(--bd)}
.perf-label{font-size:10px;font-weight:700;color:white;background:var(--a);padding:2px 8px;border-radius:var(--radius-pill);flex-shrink:0}
.perf-item{font-size:11px;color:var(--a-dark);background:white;border:1px solid var(--ab);padding:2px 9px;border-radius:var(--radius-pill);font-weight:500}

/* ── Toast & back to top ── */
.toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);background:var(--tx);color:#fff;padding:9px 20px;border-radius:var(--radius-pill);font-size:12px;font-weight:600;pointer-events:none;opacity:0;transition:all .22s;z-index:999;max-width:80vw;box-shadow:var(--shadow-md)}
.toast.on{opacity:1;transform:translateX(-50%) translateY(-4px)}
.btt{position:fixed;bottom:22px;right:22px;width:40px;height:40px;background:var(--a);color:#fff;border:none;border-radius:50%;font-size:18px;cursor:pointer;display:flex;align-items:center;justify-content:center;box-shadow:var(--shadow-md);opacity:0;pointer-events:none;transition:opacity .22s,transform .22s;z-index:998;line-height:1}
.btt.on{opacity:1;pointer-events:auto}
.btt:hover{transform:translateY(-3px);box-shadow:0 6px 18px rgba(96,125,150,.45)}
.btt:active{transform:scale(.92)}

/* ── Empty state ── */
.empty{text-align:center;padding:64px 20px;color:var(--sub)}
.empty::before{content:"";display:block;width:56px;height:56px;margin:0 auto 16px;border-radius:16px;background:linear-gradient(135deg,var(--s3),var(--s2));border:2px solid var(--bd);box-shadow:var(--shadow-xs)}

mark{background:#fef3c7;color:#92400e;border-radius:3px;padding:0 2px;font-weight:inherit}
::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:var(--bg)}::-webkit-scrollbar-thumb{background:var(--bd2);border-radius:3px}

/* ── AI Generation form ── */
.gr-wrap{max-width:840px;margin:0 auto;padding:20px 0}
.gr-card{background:#fff;border:1px solid var(--bd);border-radius:var(--radius-card);padding:0;margin-bottom:16px;overflow:hidden;box-shadow:var(--shadow-xs)}
.gr-card-head{padding:14px 18px;border-bottom:1px solid var(--bd);background:linear-gradient(to right,var(--al),var(--s));display:flex;align-items:center;gap:8px}
.gr-title{font-size:13px;font-weight:700;color:var(--a);letter-spacing:-.1px;flex:1}
.gr-card-body{padding:16px 18px}
.gr-row{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.gr-cat{padding:8px 10px;background:var(--s2);border:1.5px solid var(--bd);border-radius:8px;font-size:12px;color:var(--tx);outline:none;flex:1;transition:var(--transition)}
.gr-cat:focus{border-color:var(--a);box-shadow:0 0 0 3px rgba(96,125,150,.10)}
.gr-pct{width:64px;padding:8px 8px;background:var(--s2);border:1.5px solid var(--bd);border-radius:8px;font-size:12px;color:var(--tx);outline:none;text-align:right;transition:var(--transition)}
.gr-pct:focus{border-color:var(--a);box-shadow:0 0 0 3px rgba(96,125,150,.10)}
.gr-del{width:28px;height:28px;background:none;border:1.5px solid var(--bd2);border-radius:50%;color:var(--sub);cursor:pointer;font-size:15px;display:flex;align-items:center;justify-content:center;flex-shrink:0;line-height:1;transition:var(--transition)}
.gr-del:hover{background:#fee2e2;border-color:#fca5a5;color:#dc2626}
.gr-add{padding:7px 16px;background:var(--al);border:1.5px solid var(--ab);border-radius:var(--radius-pill);color:var(--a);font-size:12px;cursor:pointer;font-weight:600;margin-top:4px;transition:var(--transition)}
.gr-add:hover{background:var(--a);color:#fff;border-color:var(--a)}
.gr-total{font-size:12px;color:var(--sub);margin-top:10px;padding:6px 12px;background:var(--s2);border-radius:8px;border:1px solid var(--bd);display:inline-block}
.gr-total.ok{color:var(--good);font-weight:700;background:var(--goodl);border-color:var(--goodb)}
.gr-total.over{color:#dc2626;font-weight:700;background:#fee2e2;border-color:#fca5a5}
.dist-wrap{display:flex;flex-wrap:wrap;gap:6px;margin-top:4px}
.dist-cb{display:none}
.dist-lbl{padding:5px 12px;border:1.5px solid var(--bd);border-radius:var(--radius-pill);font-size:12px;color:var(--sub);cursor:pointer;transition:var(--transition)}
.dist-cb:checked+.dist-lbl{background:var(--al);border-color:var(--a);color:var(--a);font-weight:700}
.gr-poi{margin-top:10px;font-size:12px;color:var(--a);background:var(--al);border:1px solid var(--ab);border-radius:8px;padding:9px 13px;display:none;line-height:1.7}
.gr-btn{width:100%;padding:14px;background:linear-gradient(135deg,var(--a),var(--a-dark));color:#fff;border:none;border-radius:var(--radius-card);font-size:14px;font-weight:700;cursor:pointer;transition:var(--transition);letter-spacing:.2px;box-shadow:0 2px 8px rgba(96,125,150,.30);margin-top:4px}
.gr-btn:hover{background:linear-gradient(135deg,var(--a-dark),#3a5470);box-shadow:0 4px 14px rgba(96,125,150,.42);transform:translateY(-1px)}
.gr-btn:active{transform:scale(.99)}
.gr-pkg-title{display:flex;align-items:center;gap:8px;margin-bottom:10px}
.cp-all{padding:5px 14px;background:var(--s2);border:1.5px solid var(--bd2);border-radius:var(--radius-pill);font-size:11px;color:var(--sub);cursor:pointer;font-weight:600;margin-left:auto;transition:var(--transition)}
.cp-all:hover{background:var(--al);border-color:var(--ab);color:var(--a)}

/* ── Leju price panel ── */
.leju-header{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:8px}
.leju-title{font-size:11px;font-weight:600;color:var(--tx);letter-spacing:.02em}
.leju-unit{font-size:10px;color:var(--sub)}
.leju-rows{display:flex;flex-direction:column;gap:4px}
.leju-row{display:grid;grid-template-columns:46px 1fr 30px;align-items:center;gap:6px}
.leju-label{font-size:11px;color:var(--sub);text-align:right;white-space:nowrap}
.leju-label.is-current{color:var(--tx);font-weight:600}
.leju-bar-track{height:6px;background:var(--bd);border-radius:3px;overflow:hidden}
.leju-bar-fill{height:100%;border-radius:3px;background:var(--ab);transition:width .35s ease}
.leju-bar-fill.is-current{background:var(--a-dark)}
.leju-value{font-size:10px;color:var(--sub);text-align:right;white-space:nowrap}
.leju-value.is-current{color:var(--tx);font-weight:600}
.leju-source{margin-top:7px;font-size:9.5px;color:var(--bd2);text-align:right}
/* ── leju building price panel ── */
#leju-bldg-panel{margin-top:10px}
.lbp-card{background:var(--s);border:1px solid var(--bd);border-radius:var(--radius-card);box-shadow:var(--shadow-xs);padding:14px 16px 12px}
.lbp-header{display:flex;align-items:baseline;gap:10px;margin-bottom:10px}
.lbp-title{font-size:13px;font-weight:600;color:var(--tx);letter-spacing:.02em}
.lbp-subtitle{font-size:11px;color:var(--sub)}
.lbp-refline-label{font-size:11px;color:var(--a-dark);white-space:nowrap}
.lbp-chart{position:relative;padding-top:18px}
.lbp-refline{position:absolute;top:0;bottom:0;width:1px;background:var(--a);opacity:.85;pointer-events:none;z-index:2}
.lbp-refline-tick{position:absolute;top:0;transform:translateX(-50%);font-size:10px;color:var(--a-dark);font-weight:600;white-space:nowrap;background:var(--s);padding:0 3px;border-radius:3px;border:1px solid var(--ab);line-height:16px}
.lbp-row{display:grid;grid-template-columns:90px 1fr 52px;align-items:center;gap:0 7px;height:22px;margin-bottom:2px}
.lbp-row:last-child{margin-bottom:0}
.lbp-name{font-size:11px;color:var(--tx);text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:22px}
.lbp-track{position:relative;height:10px;background:var(--s3);border-radius:3px;overflow:visible}
.lbp-bar{height:100%;border-radius:3px;background:var(--ab);transition:width .35s ease}
.lbp-row--nearest .lbp-bar{background:var(--a)}
.lbp-row--nearest .lbp-name{color:var(--a-dark);font-weight:600}
.lbp-price{font-size:11px;color:var(--sub);text-align:right;white-space:nowrap;line-height:22px}
.lbp-row--nearest .lbp-price{color:var(--a-dark);font-weight:600}
.lbp-source{margin-top:8px;font-size:10px;color:var(--sub);text-align:right}
.lbp-mode-btn{margin-left:auto;padding:2px 8px;font-size:10px;background:var(--s2);border:1px solid var(--bd);border-radius:4px;cursor:pointer;color:var(--sub);font-family:inherit;line-height:1.4;flex-shrink:0}
.lbp-mode-btn:hover{color:var(--a);border-color:var(--a)}
/* ── building name autocomplete ── */
.bldg-search-block{display:flex;flex-direction:column;gap:6px;margin-bottom:8px}
.bldg-input-row{display:flex;gap:6px;align-items:stretch}
.ac-wrap{position:relative;flex:1;min-width:0}
.bldg-input{width:100%;box-sizing:border-box;padding:8px 12px;font-size:12px;color:var(--tx);background:var(--s2);border:1.5px solid var(--bd);border-radius:8px;outline:none;font-family:inherit;transition:var(--transition)}
.bldg-input:focus{border-color:var(--a)}
.bldg-input::placeholder{color:var(--sub);opacity:.7}
.ac-dropdown{position:fixed;z-index:9999;margin:0;padding:4px 0;list-style:none;background:var(--s);border:1px solid var(--bd);border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.14);max-height:218px;overflow-y:auto}
.ac-dropdown[hidden]{display:none}
.ac-item{display:flex;flex-direction:column;gap:2px;padding:9px 13px;cursor:pointer;border-left:3px solid transparent;transition:background .1s,border-color .1s}
.ac-item:hover,.ac-item[aria-selected=true]{background:var(--s2);border-left-color:var(--a)}
.ac-item__name{font-size:12px;font-weight:500;color:var(--tx)}
.ac-item:hover .ac-item__name{color:var(--a)}
.ac-item__meta{font-size:11px;color:var(--sub)}
.ac-item--loading,.ac-item--empty{flex-direction:row;align-items:center;gap:8px;cursor:default;font-size:11px;color:var(--sub)}
.ac-item--loading:hover,.ac-item--empty:hover{background:transparent;border-left-color:transparent}
.ac-spinner{display:inline-block;width:12px;height:12px;border:2px solid var(--ab);border-top-color:var(--a);border-radius:50%;flex-shrink:0;animation:ac-spin .7s linear infinite}
@keyframes ac-spin{to{transform:rotate(360deg)}}
.sec-divider{display:flex;align-items:center;gap:8px;margin:10px 0 8px;color:var(--sub);font-size:11px;user-select:none}
.sec-divider::before,.sec-divider::after{content:'';flex:1;height:1px;background:var(--bd)}
/* ── extra tags (off-database) grid redesign ── */
.xtag-section{margin-top:18px;padding-top:14px;border-top:1px solid var(--bd)}
.xtag-hd{display:flex;align-items:center;gap:8px;margin-bottom:12px}
.xtag-hd-title{font-size:11px;font-weight:700;color:var(--sub);letter-spacing:.6px;flex:1}
.xtag-shuffle-all{background:none;border:1.5px solid var(--bd);border-radius:20px;padding:3px 10px;font-size:11px;cursor:pointer;color:var(--sub);display:flex;align-items:center;gap:4px;transition:border-color .15s,color .15s;font-family:inherit}
.xtag-shuffle-all:hover{border-color:var(--a);color:var(--a)}
.xtag-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:10px}
.xtag-card{background:var(--s2);border:1px solid var(--bd);border-radius:10px;padding:10px 12px}
.xtag-card-hd{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:8px}
.xtag-card-cat{font-size:12px;font-weight:700;color:var(--tx);line-height:1.3}
.xtag-card-en{font-size:10px;color:var(--sub);margin-top:1px}
.xtag-card-shuf{background:none;border:none;cursor:pointer;font-size:15px;color:var(--sub);padding:2px;line-height:1;transition:color .15s,transform .3s;flex-shrink:0}
.xtag-card-shuf:hover{color:var(--a)}
.xtag-chips{display:flex;flex-wrap:wrap;gap:5px}
.xtag-chip{padding:4px 9px;font-size:10px;background:var(--s);border:1px solid var(--bd);border-radius:14px;cursor:pointer;color:var(--tx);transition:background .1s,border-color .1s;white-space:nowrap;display:inline-flex;align-items:center;gap:4px}
.xtag-chip:hover{background:var(--al);border-color:var(--a)}
.xtag-chip.copied{background:var(--a);color:#fff;border-color:var(--a)}
.xtag-chip-en{font-weight:500;color:var(--tx)}
.xtag-chip-zh{color:var(--sub)}
.xtag-chip:hover .xtag-chip-en,.xtag-chip:hover .xtag-chip-zh{color:var(--a-dark)}
.xtag-chip.copied .xtag-chip-en,.xtag-chip.copied .xtag-chip-zh{color:#fff}
/* ── AI reasoning box ── */
.ai-reason-box{margin-bottom:14px;padding:12px 16px;background:var(--s2);border:1px solid var(--bd2);border-radius:var(--radius-card);font-size:11px;color:var(--tx);line-height:1.7}
.ai-r-hd{font-size:11px;font-weight:700;color:var(--a-dark);margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid var(--bd);letter-spacing:.3px}
.ai-r-row{display:flex;gap:8px;padding:2px 0;align-items:baseline}
.ai-r-n{flex-shrink:0;width:16px;height:16px;background:var(--a);color:#fff;border-radius:50%;font-size:9px;font-weight:800;display:inline-flex;align-items:center;justify-content:center;margin-top:1px}
.ai-r-rules{color:var(--sub)}
/* ── Generated package output cards ── */
.ai-insight{margin-bottom:16px;padding:14px 18px;background:linear-gradient(135deg,var(--al),#f8fbff);border:1.5px solid var(--ab);border-radius:var(--radius-card);font-size:12px;color:var(--a-dark);box-shadow:var(--shadow-xs);position:relative}
.ai-insight::before{content:"AI";position:absolute;top:-1px;left:16px;background:var(--a);color:#fff;font-size:9px;font-weight:800;padding:1px 7px;border-radius:0 0 5px 5px;letter-spacing:.5px}
.ai-insight b{color:var(--tx)}
.ai-insight .ai-sub{color:var(--sub);margin-top:4px;display:block;line-height:1.5}
.pkg-out{background:var(--s);border:1px solid var(--bd);border-radius:var(--radius-card);overflow:hidden;box-shadow:var(--shadow-xs);margin-bottom:14px;transition:var(--transition)}
.pkg-out:hover{border-color:var(--ab);box-shadow:var(--shadow-sm);transform:translateY(-1px)}
.pkg-out-head{background:linear-gradient(135deg,var(--s3) 0%,var(--s) 100%);border-bottom:1px solid var(--bd);padding:12px 16px;display:flex;align-items:center;gap:8px}
.pkg-out-body{padding:12px 16px}
.pkg-loc-row{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid var(--bd)}
.pkg-desc{font-size:11px;color:var(--sub);margin-bottom:10px;padding:8px 12px;background:var(--s2);border-radius:8px;border-left:3px solid var(--ab);line-height:1.6}

/* ── Multi-select dropdown ── */
.msel{position:relative}
.msel-btn{padding:7px 12px;background:var(--s2);border:1.5px solid var(--bd);border-radius:var(--radius-pill);font-size:12px;color:var(--tx);cursor:pointer;font-weight:500;white-space:nowrap;transition:var(--transition)}
.msel-btn.has{background:var(--al);border-color:var(--ab);color:var(--a);font-weight:700}
.msel-panel{position:absolute;top:calc(100% + 5px);left:0;background:#fff;border:1px solid var(--bd);border-radius:var(--radius-card);box-shadow:var(--shadow-md);min-width:260px;z-index:200;overflow:hidden}
.msel-si{width:100%;padding:9px 12px;border:none;border-bottom:1px solid var(--bd);font-size:12px;color:var(--tx);outline:none;background:#fff}
.msel-list{max-height:220px;overflow-y:auto;padding:4px}
.msel-item{display:flex;align-items:center;gap:7px;padding:6px 8px;border-radius:6px;cursor:pointer;font-size:12px;color:var(--tx);user-select:none}
.msel-item:hover{background:var(--s2)}
.msel-foot{display:flex;justify-content:flex-end;gap:6px;padding:7px 8px;border-top:1px solid var(--bd)}
.msel-foot button{padding:5px 14px;border-radius:var(--radius-pill);border:1px solid var(--bd);background:var(--s2);color:var(--tx);font-size:11px;cursor:pointer;font-weight:600}
.msel-foot button:last-child{background:var(--a);color:#fff;border-color:var(--a)}
</style><link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"><script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script></head><body>
<header>
  <div class="logo">M</div>
  <div class="htxt"><h1>META 受眾標籤工具</h1><p>點擊標籤複製搜尋詞 → 貼入 META Ads Manager</p></div>
  <div class="hstats">
    <div class="hstat"><div class="hn">$n</div><div class="hl">標籤數</div></div>
    <div class="hstat"><div class="hn">$nc</div><div class="hl">大分類</div></div>
    <div class="hstat"><div class="hn">$ny</div><div class="hl">優等受眾包</div></div>
  </div>
</header>
<div class="tabs">
  <div class="tab on" onclick="sw('i',this)">興趣標籤庫</div>
  <div class="tab" onclick="sw('y',this)">優等受眾包</div>
  <div class="tab" onclick="sw('t',this)">類型分析</div>
  <div class="tab" onclick="sw('g',this)">AI智能生成</div>
</div>
<div class="bar" id="bi">
  <div class="sw"><input id="q" placeholder="搜尋標籤（英文或中文）..." oninput="ri()"></div>
  <select id="cat" onchange="ri()">
    <option value="">全部分類</option>
    <option value="__hidden" style="color:#9a5050;font-weight:700">隱藏興趣 ($nHidden)</option>
    <option value="__rich"   style="color:#5e7a9a;font-weight:700">有錢興趣 ($nRich)</option>
    <option value="__young"  style="color:#5e7a9a;font-weight:700">年輕興趣 ($nYoung)</option>
    <option value="__re"     style="color:#5e7a9a;font-weight:700">房地產興趣 ($nRe)</option>
    <option value="__tech"   style="color:#5e7a9a;font-weight:700">科技業興趣 ($nTech)</option>
    <option value="__civil"  style="color:#5e7a9a;font-weight:700">公務員興趣 ($nCivil)</option>
    <option disabled>──────────────</option>
$catOpts
  </select>
  <div class="cnt" id="ci">請搜尋或選分類</div>
  <button class="ttu" onclick="window.scrollTo({top:0,behavior:'smooth'})" title="回到最上面">↑</button>
</div>
<div class="bar" id="by" style="display:none">
  <div class="sw"><input id="qy" placeholder="搜尋建案名稱或興趣..." oninput="ry()"></div>
  <div class="msel" id="proj-wrap">
    <button class="msel-btn" id="proj-btn" onclick="toggleProj(event)">選擇建案 ▾</button>
    <div class="msel-panel" id="proj-panel" style="display:none">
      <input class="msel-si" placeholder="搜尋建案名稱..." oninput="filterProjOpts(this.value)">
      <div class="msel-list" id="proj-list"></div>
      <div class="msel-foot"><button onclick="clearProj()">清除</button><button onclick="toggleProj(event)">確定</button></div>
    </div>
  </div>
  <div class="cnt" id="cy">共 $ny 組</div>
  <button class="ttu" onclick="window.scrollTo({top:0,behavior:'smooth'})" title="回到最上面">↑</button>
</div>
<div class="bar" id="bt" style="display:none">
  <div class="sw"><input id="qt" placeholder="搜尋建案名稱或興趣..." oninput="rt()"></div>
  <select id="lxf" onchange="rt()"><option value="">全部類型</option>
$lxOpts
  </select>
  <div class="cnt" id="ct2">共 $ny 組</div>
  <button class="ttu" onclick="window.scrollTo({top:0,behavior:'smooth'})" title="回到最上面">↑</button>
</div>
<div class="bar" id="bg" style="display:none">
  <span style="font-size:13px;font-weight:600;color:var(--tx)">智能生成受眾包</span>
  <button class="ttu" onclick="window.scrollTo({top:0,behavior:'smooth'})" title="回到最上面">↑</button>
</div>
<div class="main">
  <div class="pg on" id="pi"><div id="io"></div><div id="extra-tags-section"></div></div>
  <div class="pg" id="py"><div id="yo" class="pgrid"></div></div>
  <div class="pg" id="pt"><div id="to" class="pgrid"></div></div>
  <div class="pg" id="pg">
    <div class="gr-wrap">
      <div class="gr-card">
        <div class="gr-card-head"><div class="gr-title">案件條件（AI依此分析買家輪廓）</div></div>
        <div class="gr-card-body">
        <div style="display:flex;gap:16px;margin-bottom:12px;font-size:12px">
          <label style="display:flex;align-items:center;gap:5px;cursor:pointer"><input type="radio" name="ptype" value="total" checked onchange="togglePtype()"> 總價範圍</label>
          <label style="display:flex;align-items:center;gap:5px;cursor:pointer"><input type="radio" name="ptype" value="unit" onchange="togglePtype()"> 每坪單價</label>
        </div>
        <div id="price-total" style="display:flex;align-items:center;gap:6px;margin-bottom:10px;flex-wrap:wrap">
          <input id="pt1" type="number" class="gr-pct" style="width:90px" placeholder="1000" oninput="updateCriteria()"> <span style="font-size:12px;color:var(--sub)">萬</span>
          <span style="color:var(--sub)">－</span>
          <input id="pt2" type="number" class="gr-pct" style="width:90px" placeholder="1500" oninput="updateCriteria()"> <span style="font-size:12px;color:var(--sub)">萬</span>
        </div>
        <div id="price-unit" style="display:none;align-items:center;gap:6px;margin-bottom:10px">
          <input id="pu1" type="number" class="gr-pct" style="width:90px" placeholder="50" oninput="updateCriteria()"> <span style="font-size:12px;color:var(--sub)">萬／坪</span>
        </div>
        <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
          <span style="font-size:12px;color:var(--sub)">坪數：</span>
          <input id="sz1" type="number" class="gr-pct" style="width:80px" placeholder="30" oninput="updateCriteria()"> <span style="font-size:12px;color:var(--sub)">坪</span>
          <span style="color:var(--sub)">－</span>
          <input id="sz2" type="number" class="gr-pct" style="width:80px" placeholder="45" oninput="updateCriteria()"> <span style="font-size:12px;color:var(--sub)">坪</span>
        </div>
        <div style="margin-top:14px;padding-top:12px;border-top:1px solid var(--bd)">
          <div style="font-size:12px;font-weight:600;color:var(--tx);margin-bottom:8px">廣告類型</div>
          <div style="display:flex;gap:20px;font-size:12px">
            <label style="display:flex;align-items:center;gap:5px;cursor:pointer"><input type="radio" name="adtype" value="名單" checked onchange="updateCriteria()"> 名單</label>
            <label style="display:flex;align-items:center;gap:5px;cursor:pointer"><input type="radio" name="adtype" value="訊息" onchange="updateCriteria()"> 訊息</label>
            <label style="display:flex;align-items:center;gap:5px;cursor:pointer"><input type="radio" name="adtype" value="1510" onchange="updateCriteria()"> 1510</label>
          </div>
        </div>
        <div id="criteria-box" style="display:none;margin-top:12px;padding:12px 14px;background:var(--s2);border:1px solid var(--bd);border-radius:8px;font-size:11px;line-height:1.9;color:var(--sub)"></div>
        </div>
      </div>
      <div class="gr-card">
        <div class="gr-card-head"><div class="gr-title">地點設定（選填）</div></div>
        <div class="gr-card-body">
        <div style="font-size:11px;color:var(--sub);margin-bottom:10px;line-height:1.6">輸入地址定位 → AI 自動識別地區類型並影響標籤推薦權重</div>
        <div id="addr-search-block" style="display:flex;gap:6px;margin-bottom:10px">
          <input id="map-addr" type="text" placeholder="輸入地址或地標，例：台北市信義區松高路" style="flex:1;padding:8px 12px;background:var(--s2);border:1.5px solid var(--bd);border-radius:8px;font-size:12px;color:var(--tx);outline:none;transition:var(--transition)" onkeydown="if(event.key===\'Enter\')searchAddr()">
          <button onclick="clearBldg();searchAddr();" style="padding:8px 16px;background:var(--a);color:#fff;border:none;border-radius:8px;font-size:12px;cursor:pointer;font-weight:600;white-space:nowrap;flex-shrink:0;box-shadow:var(--shadow-xs);transition:var(--transition)">定位</button>
        </div>
        <div id="gr-map" style="height:240px;border-radius:8px;overflow:hidden;border:1px solid var(--bd);margin-bottom:10px"></div>
        <div id="map-info" style="font-size:11px;color:var(--sub);margin-bottom:10px;padding:8px 12px;background:var(--al);border:1px solid var(--ab);border-radius:8px;display:none;line-height:1.5"></div>
        <div id="leju-panel-anchor"></div>
        <div id="leju-bldg-panel"></div>
        <div id="bldg-analysis-anchor"></div>
        <div id="real-price-anchor"></div>
        <select id="gr-city" onchange="updCity()" style="padding:8px 12px;background:var(--s2);border:1.5px solid var(--bd);border-radius:8px;font-size:12px;color:var(--tx);outline:none;width:100%;margin-bottom:12px;transition:var(--transition)">
          <option value="">選擇縣市</option>
          <option>台北市</option><option>新北市</option><option>桃園市</option>
          <option>台中市</option><option>台南市</option><option>高雄市</option>
          <option>基隆市</option><option>新竹市</option><option>新竹縣</option>
          <option>苗栗縣</option><option>彰化縣</option><option>南投縣</option>
          <option>雲林縣</option><option>嘉義市</option><option>嘉義縣</option>
          <option>屏東縣</option><option>宜蘭縣</option><option>花蓮縣</option><option>台東縣</option>
        </select>
        <div style="font-size:11px;color:var(--sub);margin-bottom:8px;font-weight:600">選擇行政區（可複選，AI依此推薦附近地標）</div>
        <div id="dist-wrap" class="dist-wrap"><span style="font-size:12px;color:var(--sub)">請先選擇縣市</span></div>
        <div id="gr-poi" class="gr-poi"></div>
        <div style="margin-top:16px;padding-top:14px;border-top:1px solid var(--bd)">
          <div style="font-size:12px;font-weight:700;color:var(--tx);margin-bottom:4px">可接受投放地點 <span style="font-size:11px;font-weight:400;color:var(--sub)">（不限縣市，逗號或頓號分隔）</span></div>
          <div style="font-size:11px;color:var(--sub);margin-bottom:8px;line-height:1.5">建案在林口但想投南港、內湖？在這裡輸入額外地點，每組受眾包會納入</div>
          <input id="extra-locs" type="text" placeholder="例：南港、內湖、板橋、竹北" style="width:100%;padding:8px 12px;background:var(--s2);border:1.5px solid var(--bd);border-radius:8px;font-size:12px;color:var(--tx);outline:none;transition:var(--transition)" oninput="previewExtraLocs()">
          <div id="extra-poi-preview" style="display:none;margin-top:8px;font-size:11px;color:var(--a);background:var(--al);border:1px solid var(--ab);border-radius:8px;padding:8px 12px;line-height:1.8"></div>
        </div>
        </div>
      </div>
      <div class="gr-card">
        <div class="gr-card-head"><div class="gr-title">興趣分類 ＆ 占比設定</div></div>
        <div class="gr-card-body">
        <div id="gr-rows"></div>
        <button class="gr-add" onclick="addGrRow()">＋ 新增分類</button>
        <div class="gr-total" id="gr-total">合計：0%</div>
        </div>
      </div>
      <button class="gr-btn" onclick="genPkg()">AI 智能分析並生成 3 組受眾包</button>
      <div id="gr-out" style="margin-top:16px"></div>
    </div>
  </div>
</div>
<div class="toast" id="toast"></div>
<button class="btt" id="btt" onclick="window.scrollTo({top:0,behavior:'smooth'})" title="回到最上面">↑</button>
<script>
$jsD
$jsY
"@

$jsFuncs = @'
function sw(n,el){
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  el.classList.add('on');
  document.querySelectorAll('.pg').forEach(p=>p.classList.remove('on'));
  document.getElementById('p'+n).classList.add('on');
  ['bi','by','bt','bg'].forEach(id=>{const e=document.getElementById(id);if(e)e.style.display='none'});
  const bar=document.getElementById('b'+n);if(bar)bar.style.display='flex';
  if(n==='g')setTimeout(initMap,80);
}
function hl(s,q){
  if(!q||!s)return s||'';
  try{return s.replace(new RegExp('('+q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+')','gi'),'<mark>$1</mark>')}catch(e){return s}
}
function ri(){
  const q=document.getElementById('q').value.trim().toLowerCase();
  const cat=document.getElementById('cat').value;
  const out=document.getElementById('io');
  if(!q&&!cat){
    document.getElementById('ci').textContent='請搜尋或選分類';
    out.innerHTML='<div class="empty"><div style="font-size:28px;margin-bottom:12px">🔍</div><div style="font-size:14px;font-weight:600;color:#374151;margin-bottom:6px">請輸入關鍵字或選擇分類</div><div style="font-size:12px">共 '+D.length+' 筆標籤</div></div>';
    return;
  }
  const isRich=(cat==='__rich'),isYoung=(cat==='__young'),isHidden=(cat==='__hidden'),isRe=(cat==='__re'),isTech=(cat==='__tech'),isCivil=(cat==='__civil');
  const spMode=isRich||isYoung||isHidden||isRe||isTech||isCivil;
  const f=D.filter(d=>{
    const cOk=!cat||(isRich?d.sp==='rich':isYoung?d.sp==='young':isHidden?d.sp==='hidden':isRe?d.sp==='re':isTech?d.sp==='tech':isCivil?d.sp==='civil':d.c===cat);
    const qOk=!q||d.cp.toLowerCase().includes(q)||(d.z&&d.z.toLowerCase().includes(q))||d.c.toLowerCase().includes(q);
    return cOk&&qOk;
  });
  document.getElementById('ci').textContent='共 '+f.length+' 筆';
  if(!f.length){out.innerHTML='<div class="empty">沒有符合的標籤</div>';return}
  const LIMIT=300,show=f.slice(0,LIMIT),g={};
  show.forEach(d=>{if(!g[d.c])g[d.c]=[];g[d.c].push(d)});
  const more=f.length>LIMIT?'<div style="text-align:center;padding:12px;font-size:11px;color:var(--sub);background:var(--s2);border-radius:6px;border:1px solid var(--bd);margin-top:8px">顯示前 '+LIMIT+' 筆，共 '+f.length+' 筆，請縮小搜尋範圍</div>':'';
  out.innerHTML=Object.entries(g).sort((a,b)=>b[1].length-a[1].length).map(([c,items])=>
    '<div class="sec"><div class="sh"><div class="st">'+c+'</div><div class="sc">'+items.length+' 個</div><div class="ln"></div></div><div class="grid">'+
    items.map(d=>{
      const spBadge=d.sp==='rich'?'<div class="tc tc-rich">有錢興趣</div>':d.sp==='young'?'<div class="tc tc-young">年輕興趣</div>':d.sp==='hidden'?'<div class="tc tc-hidden">隱藏興趣</div>':d.sp==='re'?'<div class="tc tc-re">房地產興趣</div>':d.sp==='tech'?'<div class="tc tc-tech">科技業興趣</div>':d.sp==='civil'?'<div class="tc tc-civil">公務員興趣</div>':'';
      const zhCat=d.c&&d.c!=='其他'?(d.z?d.z+'/'+d.c:d.c):(d.z||'');
      return '<div class="card" onclick="cp(this)"><span class="cb">複製</span><div class="te" data-cp="'+d.cp.replace(/"/g,'&quot;')+'">'+hl(d.cp,q)+'</div>'+(zhCat?'<div class="tz">'+hl(zhCat,q)+'</div>':'')+spBadge+'</div>';
    }).join('')+'</div></div>'
  ).join('')+more;
}
function pkgHtml(d,q){
  const tags=d.tags?d.tags.split('||').map(e=>{const p=e.indexOf('::');return p>-1?{t:e.substring(0,p),c:e.substring(p+2)}:{t:e,c:''};}).filter(x=>x.t.length>1):[];
  const thtml=tags.map(x=>'<div class="ptag" onclick="ct(\''+x.t.replace(/\\/g,'\\\\').replace(/'/g,"\\'")+'\')">'+hl(x.t,q)+(x.c&&x.c!=='其他'?'<span class="ptag-c">'+x.c+'</span>':'')+' </div>').join('');
  const locHtml=d.loc?'<div class="ploc">'+d.loc+'</div>':'';
  const lxHtml=d.lx?'<span style="font-size:10px;color:var(--a);background:var(--al);border:1px solid var(--ab);padding:2px 6px;border-radius:5px;margin-left:6px;font-weight:600">'+d.lx+'</span>':'';
  return '<div class="pkg"><div class="phd"><div class="pbadge">優</div><div style="flex:1"><div class="ptitle">'+hl(d.proj||'未命名',q)+lxHtml+'</div><div class="psub">'+hl(d.ta,q)+'</div>'+locHtml+'</div></div><div class="ptags">'+thtml+'</div></div>';
}
let selProj=new Set();
function buildProjList(){
  const all=YOU.map(d=>d.proj).filter(Boolean);
  const named=[...new Set(all.filter(p=>!p.startsWith('AI識別：')))].sort();
  const aiCnt=all.filter(p=>p.startsWith('AI識別：')).length;
  const list=document.getElementById('proj-list');
  if(!list)return;
  let h='';
  if(aiCnt>0)h='<label class="msel-item" style="color:var(--sub);border-bottom:1px solid var(--bd);margin-bottom:4px;padding-bottom:6px"><input type="checkbox" class="proj-cb" value="__ai" onchange="onProjCb(this)"> AI識別 (共'+aiCnt+'筆)</label>';
  h+=named.map(p=>'<label class="msel-item"><input type="checkbox" class="proj-cb" value="'+p.replace(/"/g,'&quot;')+'" onchange="onProjCb(this)"> '+p+'</label>').join('');
  list.innerHTML=h;
}
function toggleProj(e){if(e)e.stopPropagation();const p=document.getElementById('proj-panel');p.style.display=p.style.display==='none'?'block':'none';}
function onProjCb(el){
  if(el.checked)selProj.add(el.value);else selProj.delete(el.value);
  const btn=document.getElementById('proj-btn');
  btn.textContent=selProj.size?'已選 '+selProj.size+' 個建案 ▾':'選擇建案 ▾';
  btn.className='msel-btn'+(selProj.size?' has':'');
  ry();rt();
}
function clearProj(){
  selProj.clear();
  document.querySelectorAll('.proj-cb').forEach(c=>c.checked=false);
  document.getElementById('proj-btn').textContent='選擇建案 ▾';
  document.getElementById('proj-btn').className='msel-btn';
  ry();rt();
}
function filterProjOpts(q){document.querySelectorAll('#proj-list .msel-item').forEach(el=>{el.style.display=el.textContent.toLowerCase().includes(q.toLowerCase())?'':'none';});}
document.addEventListener('click',function(e){const w=document.getElementById('proj-wrap');if(w&&!w.contains(e.target)){const p=document.getElementById('proj-panel');if(p)p.style.display='none';}});
function ry(){
  const q=document.getElementById('qy').value.trim().toLowerCase();
  const f=YOU.filter(d=>{if(selProj.size){const ok=selProj.has(d.proj)||(selProj.has('__ai')&&d.proj&&d.proj.startsWith('AI識別：'));if(!ok)return false;}return !q||d.ta.toLowerCase().includes(q)||d.proj.toLowerCase().includes(q)||d.tags.toLowerCase().includes(q);});
  // Show match count only when a query is active
  document.getElementById('cy').textContent=q?'符合 '+f.length+' 組（共 '+YOU.length+' 組）':'共 '+f.length+' 組';
  const out=document.getElementById('yo');
  if(!f.length){out.innerHTML='<div class="empty">沒有符合的受眾包</div>';return}
  // Pass q to pkgHtml so hl() can highlight matched text in results
  out.innerHTML=f.map(d=>pkgHtml(d,q)).join('');
}
function rt(){
  const q=document.getElementById('qt').value.trim().toLowerCase();
  const lx=document.getElementById('lxf').value;
  const f=YOU.filter(d=>{if(selProj.size){const ok=selProj.has(d.proj)||(selProj.has('__ai')&&d.proj&&d.proj.startsWith('AI識別：'));if(!ok)return false;}return(!lx||d.lx===lx)&&(!q||d.ta.toLowerCase().includes(q)||d.proj.toLowerCase().includes(q)||d.tags.toLowerCase().includes(q));});
  document.getElementById('ct2').textContent='共 '+f.length+' 組';
  const out=document.getElementById('to');
  if(!f.length){out.innerHTML='<div class="empty">沒有符合的受眾包</div>';return}
  if(lx){out.innerHTML=f.map(d=>pkgHtml(d,q)).join('');return}
  const g={};
  f.forEach(d=>{const k=d.lx||'未分類';if(!g[k])g[k]=[];g[k].push(d)});
  out.innerHTML=Object.entries(g).sort((a,b)=>b[1].length-a[1].length).map(([lxk,items])=>
    '<div class="sec"><div class="sh"><div class="st">'+lxk+'</div><div class="sc">'+items.length+' 組</div><div class="ln"></div></div><div class="pgrid">'+
    items.map(d=>pkgHtml(d,q)).join('')+'</div></div>'
  ).join('');
}
window.addEventListener('scroll',()=>{const b=document.getElementById('btt');b.classList.toggle('on',window.scrollY>200)});
function cp(el){ct(el.querySelector('[data-cp]').getAttribute('data-cp'))}
function ct(text){
  const done=()=>{const t=document.getElementById('toast');t.textContent='已複製 → '+text;t.classList.add('on');clearTimeout(t._x);t._x=setTimeout(()=>t.classList.remove('on'),2000)};
  if(navigator.clipboard){navigator.clipboard.writeText(text).then(done).catch(()=>fb(text,done))}else{fb(text,done)}
}
function fb(text,cb){const a=document.createElement('textarea');a.value=text;a.style.cssText='position:fixed;opacity:0';document.body.appendChild(a);a.focus();a.select();try{document.execCommand('copy')}catch(e){}document.body.removeChild(a);cb()}
const TW={
'台北市':['中正','大同','中山','松山','大安','萬華','信義','士林','北投','內湖','南港','文山','內湖科技園區','南港軟體園區'],
'新北市':['板橋','三重','中和','永和','新莊','新店','土城','蘆洲','樹林','淡水','汐止','三峽','鶯歌','林口工業區','新北產業園區','安坑工業區'],
'桃園市':['桃園','中壢','大溪','楊梅','蘆竹','大園','龜山','八德','龍潭','平鎮','桃園航空城','中壢工業區','桃園工業區','龍潭科技園區','楊梅工業區'],
'台中市':['中區','東區','南區','西區','北區','北屯','西屯','南屯','太平','大里','霧峰','烏日','豐原','潭子','大雅','神岡','沙鹿','龍井','梧棲','清水','大甲','中科園區','豐洲科技工業園區','后里科技工業園區','梧棲工業區','台中工業區'],
'台南市':['中西','東','南','北','安平','安南','永康','歸仁','新化','仁德','新營','麻豆','佳里','善化','南科園區','台南科技工業區','安平工業區','永康工業區'],
'高雄市':['新興','前金','苓雅','鹽埕','鼓山','旗津','前鎮','三民','楠梓','小港','左營','仁武','岡山','鳳山','大寮','林園','旗山','美濃','路竹科技工業區','仁武工業區','大發工業區','臨海工業區','岡山工業區'],
'基隆市':['仁愛','信義','中山','中正','七堵','暖暖','安樂','基隆工業區'],
'新竹市':['東','北','香山','竹科（新竹科學工業園區）'],
'新竹縣':['竹北','竹東','新埔','關西','湖口','新豐','湖口工業區','台元科技園區'],
'苗栗縣':['苗栗','頭份','竹南','通霄','苑裡','竹南科技工業區','頭份工業區'],
'彰化縣':['彰化','員林','和美','鹿港','溪湖','彰化工業區','全興工業區','彰濱工業區'],
'南投縣':['南投','草屯','埔里','集集','竹山','南投工業區'],
'雲林縣':['斗六','斗南','虎尾','西螺','雲林科技工業區','六輕麥寮工業區'],
'嘉義市':['東','西'],
'嘉義縣':['太保','朴子','大林','民雄','嘉義科學工業園區','嘉義工業區'],
'屏東縣':['屏東','潮州','東港','恆春','屏東科技工業區','農業生技園區'],
'宜蘭縣':['宜蘭','羅東','蘇澳','頭城','礁溪','利澤工業區'],
'花蓮縣':['花蓮','鳳林','玉里','吉安','花蓮工業區'],
'台東縣':['台東','成功','關山','台東工業區']};
const POI={'信義':['台北101','新光三越信義A8/A11','微風廣場'],'大安':['忠孝SOGO','東區商圈','頂好超市'],'中山':['晶華酒店商圈','新光三越中山'],'士林':['故宮博物院','士林夜市'],'北投':['北投溫泉區','新北投捷運站'],'內湖':['IKEA內湖','大潤發內湖'],'南港':['南港展覽館','南港車站'],'文山':['貓空纜車','政大商圈'],'內湖科技園區':['內湖科技園區管理中心','台灣大哥大總部','緯創資通'],'南港軟體園區':['南港軟體工業園區','中研院','南港展覽館2館'],'板橋':['板橋大遠百','環球購物中心'],'新店':['裕隆城購物中心','碧潭'],'淡水':['淡水老街','漁人碼頭'],'三重':['好市多三重','三重商圈'],'林口工業區':['林口工業區','林口三井outlet'],'安坑工業區':['安坑工業區','新店大豐路商圈'],'桃園航空城':['桃園國際機場','諾富特機場酒店'],'中壢工業區':['中壢工業區','中壢SOGO'],'桃園工業區':['桃園工業區','桃園高鐵站'],'龍潭科技園區':['龍潭科技園區','龍潭大池'],'南屯':['好市多南屯','廣三SOGO','文心森林公園'],'西屯':['逢甲夜市','台中歌劇院','大遠百台中'],'北屯':['新天地購物中心','北屯路商圈'],'太平':['長億商圈','太平大買家'],'烏日':['台中高鐵站','好市多烏日'],'豐原':['廟東夜市','豐原火車站'],'大里':['大里大買家','大里藝術館'],'霧峰':['亞洲大學現代美術館','霧峰農會酒莊'],'中科園區':['中科管理局','聯詠科技台中廠','台積電中科廠'],'豐洲科技工業園區':['豐洲科技工業園區','全球第三大TFT-LCD聚落'],'后里科技工業園區':['后里科技工業園區','后里馬場'],'梧棲工業區':['梧棲工業區','台中港'],'台中工業區':['台中工業區','大里都心商圈'],'南科園區':['台積電南科廠','南科管理局','奇美博物館'],'台南科技工業區':['台南科技工業區','台南高鐵站'],'中西':['赤崁樓','台南孔廟商圈','花園夜市'],'永康':['台南高鐵站','家樂福永康'],'安平':['安平老街','夕遊出張所'],'路竹科技工業區':['路竹科技工業區','橋頭糖廠'],'仁武工業區':['仁武工業區','澄清湖'],'大發工業區':['大發工業區','大寮棒球場'],'臨海工業區':['臨海工業區','高雄港'],'岡山工業區':['岡山工業區','岡山眷村博物館'],'苓雅':['夢時代購物中心','瑞豐夜市'],'前鎮':['高雄展覽館','好市多前鎮'],'左營':['漢神巨蛋','高雄高鐵站','蓮池潭'],'鳳山':['衛武營藝術文化中心','大東文化藝術中心'],'楠梓':['好市多楠梓','統一夢時代楠梓'],'竹科（新竹科學工業園區）':['台積電竹科總部','聯發科技','工研院'],'台元科技園區':['台元科技園區','竹北高鐵站'],'湖口工業區':['湖口工業區','湖口老街'],'竹南科技工業區':['竹南科技工業區','頭份交流道'],'彰濱工業區':['彰濱工業區','鹿港老街'],'彰化工業區':['彰化工業區','彰化火車站'],'雲林科技工業區':['雲林科技工業區','雲林斗六夜市'],'六輕麥寮工業區':['台塑麥寮六輕','麥寮港'],'嘉義科學工業園區':['嘉義科學工業園區','嘉義高鐵站'],'屏東科技工業區':['屏東科技工業區','屏東夜市'],'利澤工業區':['利澤工業區','羅東夜市'],'竹北':['竹北高鐵站','家樂福竹北'],'中壢':['中壢SOGO','老街溪河廊'],'桃園':['桃園高鐵站','桃園展演中心']};
function updCity(){
  const city=document.getElementById('gr-city').value;
  const dw=document.getElementById('dist-wrap');
  if(!city||!TW[city]){dw.innerHTML='<span style="font-size:12px;color:var(--sub)">請先選擇縣市</span>';updPOI();return;}
  dw.innerHTML=TW[city].map(d=>'<input type="checkbox" class="dist-cb" id="dc-'+d+'" value="'+d+'" onchange="updPOI()"><label class="dist-lbl" for="dc-'+d+'">'+d+'</label>').join('');
  updPOI();
}
function updPOI(){
  const dists=[...document.querySelectorAll('.dist-cb:checked')].map(c=>c.value);
  let pois=[];dists.forEach(d=>{if(POI[d])pois=pois.concat(POI[d])});
  const el=document.getElementById('gr-poi');
  if(pois.length){el.style.display='block';el.innerHTML='<b>AI推薦地標：</b>'+pois.join('、')+'<span style="font-size:10px;color:var(--sub);margin-left:6px">（可加入受眾包地點欄位）</span>';}
  else el.style.display='none';
}
function addGrRow(){
  const cats=[
    {v:'__auto',l:'不指定（AI智能分配）'},
    {v:'',l:'─────────',d:true},
    {v:'__hidden',l:'隱藏興趣（全部）'},{v:'__hidden_family',l:'隱藏 - 親子/家庭'},{v:'__hidden_outdoor',l:'隱藏 - 戶外/休閒'},
    {v:'__hidden_food',l:'隱藏 - 美食/生活'},{v:'__hidden_biz',l:'隱藏 - 職場/商務'},{v:'__hidden_fashion',l:'隱藏 - 時尚/美妝'},{v:'__hidden_travel',l:'隱藏 - 旅遊/風格'},
    {v:'__rich',l:'有錢興趣'},{v:'__young',l:'年輕興趣'},
    {v:'__re',l:'房地產興趣'},{v:'__tech',l:'科技業興趣'},{v:'__civil',l:'公務員興趣'},
    {v:'',l:'─────────',d:true},
    {v:'科技',l:'科技'},{v:'金融',l:'金融'},{v:'商業',l:'商業'},{v:'美食',l:'美食'},
    {v:'家居',l:'家居'},{v:'時尚美妝',l:'時尚美妝'},{v:'健康運動',l:'健康運動'},
    {v:'娛樂',l:'娛樂'},{v:'旅遊生活',l:'旅遊生活'},{v:'教育',l:'教育'},
    {v:'農業自然',l:'農業自然'},{v:'社會',l:'社會'},{v:'交通',l:'交通'},{v:'其他',l:'其他'},
    {v:'',l:'── 延伸興趣（未收錄）──',d:true},
    {v:'__xfin',l:'延伸・財務投資'},{v:'__xprop',l:'延伸・房產直接'},
    {v:'__xfam',l:'延伸・親子家庭'},{v:'__xhlt',l:'延伸・健康運動'},
    {v:'__xtech',l:'延伸・科技職場'},{v:'__xcar',l:'延伸・汽車交通'},
    {v:'__xtrip',l:'延伸・旅遊度假'},{v:'__xfood',l:'延伸・美食生活'},
    {v:'__xlux',l:'延伸・精品奢華'},{v:'__xbiz',l:'延伸・職場創業'}
  ];
  const id='r'+Date.now();
  const div=document.createElement('div');
  div.className='gr-row';div.id=id;
  div.innerHTML='<select class="gr-cat"><option value="">選擇分類</option>'+cats.map(c=>c.d?'<option disabled>'+c.l+'</option>':'<option value="'+c.v+'">'+c.l+'</option>').join('')+'</select><input class="gr-pct" type="number" min="1" max="100" placeholder="0" oninput="updTotal()"><span style="font-size:12px;color:var(--sub);flex-shrink:0">%</span><button class="gr-del" onclick="this.closest(\'.gr-row\').remove();updTotal()">×</button>';
  document.getElementById('gr-rows').appendChild(div);
  updTotal();
}
function updTotal(){
  const rows=[...document.querySelectorAll('#gr-rows .gr-pct')];
  const total=rows.reduce((s,i)=>s+(parseInt(i.value)||0),0);
  const hasAuto=[...document.querySelectorAll('#gr-rows .gr-cat')].some(s=>s.value==='__auto');
  const el=document.getElementById('gr-total');
  if(total>100){
    const over=total-100;
    el.textContent='超出 '+over+'%（請減少各項百分比，合計不得超過 100%）';
    el.className='gr-total over';
  } else if(total===100){
    el.textContent='已設定 100% ✓（所有標籤由您指定分類）';
    el.className='gr-total ok';
  } else if(hasAuto){
    const remain=100-total;
    el.textContent='已設定 '+total+'% ｜ 剩餘 '+remain+'%（AI智能補充）';
    el.className='gr-total ok';
  } else {
    el.textContent='合計：'+total+'%'+(total===0?'（請輸入百分比）':'');
    el.className='gr-total';
  }
}
function cpAll(btn){
  const tags=btn.closest('.pkg').dataset.tags.split('||').join(', ');
  ct(tags);
}
const hiddenSub={
  family:['Household','LifeStyle Home','Motherhood','Parenting','Family travel','Kids activities','Cultural','Kitchen appliances','Design studios','House work','HVAC','Cabinetry','Garden furniture','Gazebo','Tiny House Movement','KitchenAid','Home and garden'],
  outdoor:['Recreational fishing','Extreme sports','Team sports','Fitness professional','Trails','Birdwatching','Outdoor living','Hiking equipment','Top Gear','Family car','Enterprise Rent-A-Car','Rural','Longlining','Semi-truck','Forklift'],
  food:['Foodie','Healthy cooking','Local events','KIND Healthy Snacks','Food trucks','Tea leaf','Control system','Foodservice','LifeStyle Food','Healthy','Nutritional supplements','BACK','Nutrition','Protein','Healthy recipes','FineDiningLovers'],
  biz:['Payment service provider','Employee engagement','Medical Sales Representative','Freight transport','Customer support','Trade shows','Lending','Foreign company'],
  fashion:['Weekend','Fragrance oil','HydraFacial','Fashion Tech','Jewelry','Collectible','Digital media','Boutique'],
  travel:['Lifestyle brand','Lifestyle Travel','Lifestyle content','Weekend Trips','Christmas Holiday','Self-guided tour','Brunch','Happy hour','Skyscanner','Cheapflights','Video game culture']
};
var _xCatIdx={__xfin:0,__xprop:1,__xfam:2,__xhlt:3,__xtech:4,__xcar:5,__xtrip:6,__xfood:7,__xlux:8,__xbiz:9};
function spPool(cat,used){
  const spMap={__rich:'rich',__young:'young',__hidden:'hidden',__re:'re',__tech:'tech',__civil:'civil'};
  if(spMap[cat])return D.filter(d=>d.sp===spMap[cat]&&!used.has(d.cp));
  if(cat.startsWith('__hidden_')){
    const sub=cat.replace('__hidden_','');
    const list=hiddenSub[sub]||[];
    return D.filter(d=>d.sp==='hidden'&&list.includes(d.cp)&&!used.has(d.cp));
  }
  if(cat in _xCatIdx){
    var g=EXTRA_TAGS[_xCatIdx[cat]];
    if(!g)return[];
    return g.tags.filter(function(t){return!used.has(t.en);}).map(function(t){return{cp:t.en,sc:5,sp:'',c:cat};});
  }
  // For plain category pools: sort tags that have any sp value before plain tags
  // (sp-tagged items are hand-curated "special" entries within a category and rank higher)
  return D.filter(d=>d.c===cat&&!used.has(d.cp)).sort((a,b)=>(b.sp?1:0)-(a.sp?1:0)||a.cp.localeCompare(b.cp));
}
let _map=null,_marker=null;
// geoDist: Haversine geodesic distance in km between two lat/lng points
function geoDist(a,b,c,d){const R=6371,dLat=(c-a)*Math.PI/180,dLon=(d-b)*Math.PI/180,x=Math.sin(dLat/2)**2+Math.cos(a*Math.PI/180)*Math.cos(c*Math.PI/180)*Math.sin(dLon/2)**2;return R*2*Math.atan2(Math.sqrt(x),Math.sqrt(1-x));}
function detectAreaType(lat,lng){
  const tech=[
    {lat:24.79,lng:121.0,r:4,name:'新竹科學園區'},{lat:24.12,lng:120.67,r:5,name:'中部科學園區'},
    {lat:23.15,lng:120.34,r:5,name:'台南科學園區'},{lat:22.82,lng:120.42,r:5,name:'南部科學園區'},
    {lat:25.04,lng:121.62,r:2.5,name:'南港/內湖科技園區'},{lat:25.08,lng:121.58,r:2,name:'內湖科技園區'},
    {lat:24.96,lng:121.22,r:4,name:'桃園航空城/科技'}
  ];
  const luxury=[
    {lat:25.033,lng:121.565,r:1.8,name:'台北信義區'},{lat:25.026,lng:121.543,r:2,name:'台北大安區'},
    {lat:25.047,lng:121.529,r:1.5,name:'台北中山區'},{lat:24.155,lng:120.685,r:2,name:'台中七期重劃區'},
    {lat:22.616,lng:120.305,r:2,name:'高雄左營/鼓山'},{lat:22.996,lng:120.212,r:2,name:'台南安平/新市政'}
  ];
  for(const p of tech){if(geoDist(lat,lng,p.lat,p.lng)<p.r)return{type:'tech',name:p.name};}
  for(const z of luxury){if(geoDist(lat,lng,z.lat,z.lng)<z.r)return{type:'luxury',name:z.name};}
  if(lng>121.5&&lat>22.5&&lat<25)return{type:'resort',name:'東部/度假地帶'};
  if(lat<22.5)return{type:'resort',name:'南台灣/觀光地帶'};
  return{type:'general',name:''};
}
function initMap(){
  if(_map){_map.invalidateSize();return;}
  const el=document.getElementById('gr-map');
  if(!el||typeof L==='undefined')return;
  _map=L.map('gr-map').setView([23.97,120.97],7);
  window._map=_map;
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap',maxZoom:18}).addTo(_map);
  _map.on('click',function(e){
    const lat=e.latlng.lat,lng=e.latlng.lng;
    if(_marker)_marker.remove();
    _marker=L.marker([lat,lng]).addTo(_map);
    revGeocode(lat,lng);
  });
}
function searchAddr(){
  const q=document.getElementById('map-addr').value.trim();
  if(!q)return;
  const info=document.getElementById('map-info');
  if(info){info.style.display='block';info.textContent='搜尋中...';}
  fetch('https://nominatim.openstreetmap.org/search?q='+encodeURIComponent(q)+'&countrycodes=tw&format=json&limit=1')
    .then(r=>r.json()).then(function(data){
      if(!data||!data.length){if(info)info.textContent='找不到地址，請輸入更完整的地址';return;}
      const lat=parseFloat(data[0].lat),lng=parseFloat(data[0].lon);
      if(!_map)initMap();
      _map.setView([lat,lng],14);
      if(_marker)_marker.remove();
      _marker=L.marker([lat,lng]).addTo(_map);
      revGeocode(lat,lng);
    }).catch(function(){if(info)info.textContent='定位失敗，請稍後再試';});
}
function revGeocode(lat,lng){
  const info=document.getElementById('map-info');
  if(info){info.style.display='block';info.textContent='定位中...';}
  const ar=detectAreaType(lat,lng);
  window.mapAreaType=ar.type;window.mapAreaName=ar.name;
  fetch('https://nominatim.openstreetmap.org/reverse?lat='+lat+'&lon='+lng+'&format=json&accept-language=zh-TW')
    .then(r=>r.json()).then(function(data){
      const a=data.address||{};
      const cityRaw=a.city||a.county||a.state_district||'';
      const dist=a.suburb||a.district||a.town||a.village||'';
      const cityMap={'台北':'台北市','臺北':'台北市','新北':'新北市','桃園':'桃園市','台中':'台中市','臺中':'台中市',
        '台南':'台南市','臺南':'台南市','高雄':'高雄市','基隆':'基隆市','新竹市':'新竹市','新竹縣':'新竹縣',
        '苗栗':'苗栗縣','彰化':'彰化縣','南投':'南投縣','雲林':'雲林縣','嘉義市':'嘉義市','嘉義縣':'嘉義縣',
        '屏東':'屏東縣','宜蘭':'宜蘭縣','花蓮':'花蓮縣','台東':'台東縣','臺東':'台東縣'};
      let matched='';
      for(const k in cityMap){if(cityRaw.includes(k)){matched=cityMap[k];break;}}
      const sel=document.getElementById('gr-city');
      if(sel&&matched){
        for(const o of sel.options){if(o.text===matched){sel.value=o.value;updCity();break;}}
      }
      window.mapLocKey=(dist||cityRaw||'');
      showLejuPanel(matched||cityRaw,dist);
      var distKey=(dist||'').replace(/[區市鄉鎮村]$/,'');
      window.mapDistKey=distKey;
      window.mapCityKey=matched||cityRaw||'';
      showLejuBldgPanel(distKey,computeUserUnitPrice());
      showRealPricePanel(distKey,window.mapCityKey);
      let areaLabel={tech:'科技/園區地帶',luxury:'精華豪宅地帶',resort:'度假/觀光地帶',general:''}[ar.type]||'';
      let txt=(cityRaw||'未知')+(dist?' '+dist:'');
      if(ar.name)txt+=' ｜ '+ar.name;
      if(areaLabel)txt+='（'+areaLabel+'）';
      if(info)info.textContent=txt;
      if(_marker)_marker.bindPopup('<b>'+txt+'</b>').openPopup();
    }).catch(function(){if(info)info.textContent='無法取得地址（地區類型已偵測）';});
}
function togglePtype(){
  const t=document.querySelector('input[name="ptype"]:checked').value;
  document.getElementById('price-total').style.display=t==='total'?'flex':'none';
  document.getElementById('price-unit').style.display=t==='unit'?'flex':'none';
  updateCriteria();
}
function analyzeBuyer(){
  const priceType=document.querySelector('input[name="ptype"]:checked').value;
  const s1=parseFloat(document.getElementById('sz1').value)||0;
  const s2=parseFloat(document.getElementById('sz2').value)||s1;
  let totalAvg=0,unitAvg=0;
  if(priceType==='total'){
    const p1=parseFloat(document.getElementById('pt1').value)||0;
    const p2=parseFloat(document.getElementById('pt2').value)||p1;
    totalAvg=(p1+p2)/2;
    if(s1+s2>0)unitAvg=totalAvg/((s1+s2)/2);
  } else {
    unitAvg=parseFloat(document.getElementById('pu1').value)||0;
    if(s1+s2>0)totalAvg=unitAvg*((s1+s2)/2);
  }
  const avgSize=(s1+s2)/2||0;
  return {totalAvg,unitAvg,avgSize,
    isLuxury:totalAvg>3000||unitAvg>100,
    isHighEnd:totalAvg>=1500||unitAvg>=60,
    isMid:totalAvg>=800||unitAvg>=30,
    isSmall:avgSize>0&&avgSize<25,
    isLarge:avgSize>50};
}
function getStrategies(p){
  if(p.isLuxury)return[
    {name:'頂奢豪宅買家',desc:'高淨值精品生活族群，豪宅主力客群，重視品牌地段象徵與奢華居住體驗',cats:[['__rich',45],['金融',25],['旅遊生活',20],['美食',10]]},
    {name:'資產配置投資客',desc:'財富管理意識強，視不動產為資產配置工具，追求保值與高流動性增值',cats:[['金融',40],['__rich',30],['商業',20],['__hidden_biz',10]]},
    {name:'頂奢品味換屋',desc:'追求生活質感的成熟換屋族，重視社區等級、精品周邊與私密性',cats:[['__rich',40],['娛樂',20],['旅遊生活',25],['時尚美妝',15]]}
  ];
  if(p.isHighEnd)return[
    {name:'中高資產換屋族',desc:'具換屋實力的熟齡買家，重視格局空間、優質學區與成熟生活機能',cats:[['家居',25],['健康運動',20],['金融',20],['旅遊生活',20],['教育',15]]},
    {name:'理財型投資買家',desc:'看重增值潛力與穩定租金收益，財務規劃意識強的理性置產族',cats:[['金融',35],['房地產興趣',25],['商業',20],['__hidden_biz',20]]},
    {name:'生活品味追求者',desc:'重視居住體驗與質感氛圍，對社區美食文化與周邊生活資源敏感',cats:[['旅遊生活',30],['美食',25],['娛樂',25],['時尚美妝',20]]}
  ];
  if(p.isMid)return[
    {name:'首購換屋家庭',desc:'計畫購屋的三口小家庭，重視學區機能、坪效規劃與社區安全',cats:[['教育',25],['家居',20],['旅遊生活',20],['健康運動',20],['__young',15]]},
    {name:'中價位投資客',desc:'以租養貸或資產增值為目標，關注投報率與地段租賃需求強度',cats:[['金融',35],['房地產興趣',30],['商業',20],['科技',15]]},
    {name:'生活感首購族',desc:'重視日常生活感的年輕買家，對周邊機能、交通便利與社群氣氛敏感',cats:[['旅遊生活',30],['美食',25],['__young',25],['娛樂',20]]}
  ];
  return[
    {name:'年輕首購主力',desc:'社會新鮮人或新婚族，強調捷運便利、低總價輕鬆入場與未來增值空間',cats:[['__young',35],['旅遊生活',25],['娛樂',20],['科技',20]]},
    {name:'小坪數投資族',desc:'以出租獲利為主，鎖定學區或工作機能強烈的精華地段小宅',cats:[['金融',40],['商業',30],['__hidden_biz',20],['科技',10]]},
    {name:'實用機能族群',desc:'注重CP值與日常生活便利，以機能完善、公設實用為首要考量',cats:[['美食',30],['健康運動',30],['旅遊生活',20],['教育',20]]}
  ];
}
// Keyword sets for real-estate buyer relevance scoring within each special category.
// Tags with direct purchase-intent language rank higher than tangentially-related ones.
const _richKw  =['Luxury','Premium','Wealth','Private','Exclusive','High-net-worth','Fine dining','Private bank','Asset management','Boutique hotel','First class','Hedge fund','Private equity','Certified Financial'];
const _youngKw =['TikTok','Instagram','Esport','Streaming','K-pop','Pop culture','Anime','Sneaker','Music festival','Social media','Cryptocurrency','Startup','NFT','Mobile app','Podcast'];
const _reKw    =['Real estate','Mortgage','Home buyer','Condominium','Townhouse','Residential','Renovation','Open house','Rental','Real estate developer'];
const _techKw  =['Software','DevOps','Machine learning','Artificial intelligence','Cloud','Cybersecurity','Open source','Programming','Data science','GitHub','Stack Overflow','Android'];
const _civilKw =['Retirement','Pension','Life insurance','Public service','Government','Civil service','Cultural tourism','Art exhibition','Continuing education','Non-profit'];
function _kwBonus(name,kwList){
  // Returns 0..2 — 1 point per keyword hit, capped at 2
  let hits=0;
  for(let i=0;i<kwList.length&&hits<2;i++){if(name.toLowerCase().includes(kwList[i].toLowerCase()))hits++;}
  return hits;
}
function scoreTag(d,cat){
  let base;
  if(cat==='__rich'&&d.sp==='rich'){base=10+_kwBonus(d.cp,_richKw);}
  else if(cat==='__young'&&d.sp==='young'){base=10+_kwBonus(d.cp,_youngKw);}
  else if(cat==='__hidden'&&d.sp==='hidden'){base=10;}
  else if((cat==='房地產興趣'||cat==='__re')&&d.sp==='re'){base=8+_kwBonus(d.cp,_reKw);}
  else if((cat==='科技'||cat==='__tech')&&d.sp==='tech'){base=8+_kwBonus(d.cp,_techKw);}
  else if(cat==='__civil'&&d.sp==='civil'){base=8+_kwBonus(d.cp,_civilKw);}
  else{
    const at=window.mapAreaType||'general';
    if(at==='tech'&&(d.sp==='tech'||d.c==='科技')){base=7;}
    else if(at==='luxury'&&d.sp==='rich'){base=7;}
    else if(at==='resort'&&(d.c==='旅遊生活'||d.c==='娛樂')){base=5;}
    else if(d.c===cat){
      const catKwMap={'房地產興趣':_reKw,'科技':_techKw,'金融':_richKw};
      const kw=catKwMap[cat];
      base=kw?2+_kwBonus(d.cp,kw):2;
    }else{base=1;}
  }
  const mCatMap={'__rich':'__rich','__young':'__family','__hidden':'__rich',
    '房地產興趣':'__re','__re':'__re','科技':'__tech','__tech':'__tech','__civil':'__civil'};
  const mCat=mCatMap[cat]||'__general';
  const unitPrice=window.lejuBldgPrice||computeUserUnitPrice()||55;
  return Math.round(base*getTagMultiplier(mCat,unitPrice,window.mapAreaType||'general'));
}
function matchExtraLocs(val){
  const parts=val.split(/[,，、\s]+/).map(s=>s.trim()).filter(Boolean);
  let pois=[];
  parts.forEach(function(d){
    if(POI[d]){pois=pois.concat(POI[d]);}
    else{Object.keys(POI).forEach(function(k){if(k.includes(d)||(d.length>=2&&k.startsWith(d)))pois=pois.concat(POI[k]);});}
  });
  return [...new Set(pois)];
}
function previewExtraLocs(){
  const val=(document.getElementById('extra-locs')||{}).value||'';
  const pois=matchExtraLocs(val);
  const el=document.getElementById('extra-poi-preview');
  if(!el)return;
  if(pois.length){el.style.display='block';el.innerHTML='<b>對應地標：</b>'+pois.join('、');}
  else el.style.display='none';
}
function getExtraPois(){return matchExtraLocs((document.getElementById('extra-locs')||{}).value||'');}
function genPkg(){
  const p=analyzeBuyer();
  const hasInput=document.getElementById('pt1').value||document.getElementById('pt2').value||document.getElementById('pu1').value||document.getElementById('sz1').value;
  if(!hasInput){alert('請填入價格或坪數條件，AI 才能分析買家輪廓');return;}

  // ── Validate user-defined interest rows before generating ──
  const rows=[...document.querySelectorAll('#gr-rows .gr-row')];
  let fixedCats=[],autoPct=0;
  rows.forEach(function(r){
    const cat=r.querySelector('.gr-cat').value;
    const pct=parseInt(r.querySelector('.gr-pct').value)||0;
    if(!cat||!pct)return;
    if(cat==='__auto')autoPct+=pct;
    else fixedCats.push([cat,pct]);
  });
  const useUserCats=fixedCats.length>0||autoPct>0;
  if(useUserCats){
    const totalPct=fixedCats.reduce((s,c)=>s+c[1],0)+autoPct;
    // Warn if percentage rows exist but all are zero
    if(totalPct===0&&rows.length>0){
      alert('您已新增分類列，但所有百分比均為 0%。\n請為每個分類輸入百分比，或刪除所有列讓 AI 自動分配。');return;
    }
    // Warn and stop if total exceeds 100%
    if(totalPct>100){
      alert('各項百分比合計為 '+totalPct+'%，超出 '+(totalPct-100)+'%。\n請調整各項百分比，合計不得超過 100%。');return;
    }
  }

  const city=document.getElementById('gr-city').value||'';
  const dists=[...document.querySelectorAll('.dist-cb:checked')].map(c=>c.value);
  let basePois=[];dists.forEach(d=>{if(POI[d])basePois=basePois.concat(POI[d])});
  // Auto-fill from city if no districts selected
  if(basePois.length===0&&city&&TW[city]){
    TW[city].forEach(d=>{if(POI[d])basePois=basePois.concat(POI[d]);});
  }
  const extraPois=getExtraPois();
  // Deduplicated combined pool: base first, then extra
  const allPois=[...new Set([...basePois,...extraPois])];

  const strategies=getStrategies(p);
  const out=document.getElementById('gr-out');
  const labels=['第1組','第2組','第3組'];
  const tier=p.isLuxury?'豪宅頂奢':p.isHighEnd?'中高價位':p.isMid?'中價位':'平價首購';
  const priceStr=p.totalAvg>0?'總價約'+Math.round(p.totalAvg)+'萬':(p.unitAvg>0?'單價約'+Math.round(p.unitAvg)+'萬/坪':'');
  const sizeStr=p.avgSize>0?'坪數約'+Math.round(p.avgSize)+'坪':'';
  const specStr=p.isLarge?'大坪數':p.isSmall?'小坪數':'';
  const areaType=window.mapAreaType||'general';
  const areaName=window.mapAreaName||'';
  const areaLabel={tech:'科技/園區地帶',luxury:'精華豪宅地帶',resort:'度假/觀光地帶'}[areaType]||'';
  const catMode=!useUserCats?'AI依案件等級自動分配興趣組合'
    :autoPct>0?'依您設定占比生成，其中 '+autoPct+'% 由 AI 智能補充（'+strategies[0].name+'策略）'
    :'依您設定的興趣占比生成，AI依價格優先排序標籤';
  const parts=[priceStr,sizeStr,specStr,tier+'案件'];
  if(areaName)parts.push(areaName+(areaLabel?' ('+areaLabel+')':''));
  // ── AI reasoning section ──
  // Part B: this-case evaluation
  const p1r=p.totalAvg>0?'總價 '+Math.round(p.totalAvg)+'萬':'';
  const p2r=p.unitAvg>0?'單價 '+Math.round(p.unitAvg)+'萬/坪':'';
  const priceEvid=(p1r&&p2r)?p1r+' ／ '+p2r:(p1r||p2r||'未輸入價格');
  const tierReason=p.isLuxury?priceEvid+' → 豪宅頂奢策略':p.isHighEnd?priceEvid+' → 中高價換屋策略':p.isMid?priceEvid+' → 中價位策略':priceEvid+' → 平價首購策略';
  const areaReason=areaType==='tech'?'偵測到科技園區 → 科技業標籤獲 +7 地區加成':areaType==='luxury'?'偵測到精華豪宅地帶 → 有錢興趣標籤獲 +7 地區加成':areaType==='resort'?'偵測到度假觀光地帶 → 旅遊娛樂標籤獲 +5 地區加成':'一般住宅地帶 → 無地區加成';
  const topScoreLabel=p.isLuxury?'有錢興趣標籤（12分）':p.isHighEnd?'有錢興趣 ／ 房地產直接標籤（10分）':p.isMid?'房地產直接標籤（10分）':'房地產直接標籤（8分）';
  const stratRows=strategies.map(function(s,i){return '<div style="margin-top:3px"><b>第'+(i+1)+'組 '+s.name+'</b>：'+s.cats.map(function(c){return c[0]+' '+c[1]+'%';}).join(' ／ ')+'</div>';}).join('');
  let html='<div class="ai-reason-box">';
  html+='<div class="ai-r-hd">建案分析<\/div>';
  html+='<div class="ai-r-row"><span class="ai-r-n">1</span><span>'+tierReason+'</span></div>';
  html+='<div class="ai-r-row"><span class="ai-r-n">2</span><span>地區：'+areaReason+'</span></div>';
  html+='<div class="ai-r-row"><span class="ai-r-n">3</span><span>最高優先標籤類型：'+topScoreLabel+'</span></div>';
  html+='<div class="ai-r-row"><span class="ai-r-n">4</span><span>生成策略：'+stratRows+'</span></div>';
  html+='</div>';
  html+='<div style="margin-bottom:14px;padding:12px 16px;background:var(--al);border:1px solid var(--ab);border-radius:10px;font-size:12px;color:var(--a)"><b>AI 分析：</b>'+parts.filter(Boolean).join(' ／ ')+'<br><span style="color:var(--sub)">'+catMode+'</span></div>';

  // ── Per-package deduplication: reset used Set each package ──
  // Rationale: using a single global `used` across 3 packages means packages 2 and 3
  // are forced to use progressively lower-quality tags, since the best tags were
  // consumed by package 1. Each package targets a different buyer persona, so
  // it is correct for each to independently select the highest-ranked tags for its
  // own category mix. Overlap across packages is acceptable (ads manager handles it).
  const allTagsCombined=new Set(); // tracks all tags ever used, for the summary stat
  let pkgSelections=[];

  for(let pkg=0;pkg<3;pkg++){
    const strat=strategies[pkg];
    // Each package starts with a clean deduplication slate
    const used=new Set();
    // Determine category list: user-fixed + AI填auto / 全AI / 全用戶
    let catList,catTotal;
    if(!useUserCats){catList=strat.cats;catTotal=100;}
    else if(autoPct===0){catList=fixedCats;catTotal=fixedCats.reduce((s,c)=>s+c[1],0)||100;}
    else{
      const fixedTotal=fixedCats.reduce((s,c)=>s+c[1],0);
      const aiBase=strat.cats.filter(c=>!fixedCats.find(f=>f[0]===c[0]));
      const aiBaseTotal=aiBase.reduce((s,c)=>s+c[1],0)||1;
      const aiScaled=aiBase.map(c=>[c[0],Math.round(c[1]/aiBaseTotal*autoPct)]).filter(c=>c[1]>0);
      catList=[...fixedCats,...aiScaled];
      catTotal=(fixedTotal+autoPct)||100;
    }
    const sel=[];
    catList.forEach(function(pair){
      const cat=pair[0],pct=pair[1];
      const n=Math.max(1,Math.round(30*pct/catTotal));
      const pool=spPool(cat,used);
      pool.sort(function(a,b){return scoreTag(b,cat)-scoreTag(a,cat)||a.cp.localeCompare(b.cp);});
      pool.slice(0,n).forEach(function(d){used.add(d.cp);sel.push(d);});
    });
    // Fill to 30 using top category
    if(sel.length<30){
      const fillCat=catList[0][0];
      spPool(fillCat,used).sort(function(a,b){return scoreTag(b,fillCat)-scoreTag(a,fillCat)||a.cp.localeCompare(b.cp);}).slice(0,30-sel.length).forEach(function(d){used.add(d.cp);sel.push(d);});
    }
    if(sel.length<30){
      D.filter(function(d){return !used.has(d.cp);}).sort(function(a,b){return a.cp.localeCompare(b.cp);}).slice(0,30-sel.length).forEach(function(d){used.add(d.cp);sel.push(d);});
    }
    const s30=sel.slice(0,30);
    s30.forEach(function(d){allTagsCombined.add(d.cp);});
    pkgSelections.push(s30);
  }

  // Summary stat: total unique tags across all 3 packages (after dedup)
  html+='<div style="margin-bottom:12px;padding:8px 14px;background:var(--goodl);border:1px solid var(--goodb);border-radius:8px;font-size:12px;color:var(--good);font-weight:600">已生成，共選出 '+allTagsCombined.size+' 個標籤</div>';

  for(let pkg=0;pkg<3;pkg++){
    const strat=strategies[pkg];
    const s30=pkgSelections[pkg];
    const tagIds=s30.map(function(d){return d.cp;}).join('||');
    const pkgPois=allPois.length?[0,1,2].map(function(i){return allPois[(pkg*3+i)%allPois.length];}).filter(function(v,i,a){return a.indexOf(v)===i;}).slice(0,Math.min(3,allPois.length)):[];
    // radii: display radius rings (km) for the up-to-3 POI pins in each package
    const radii=[2,3,5];
    const locHtml=pkgPois.length?'<div style="display:flex;flex-wrap:wrap;gap:4px;margin-bottom:8px">'+pkgPois.map(function(pp,i){return '<span style="font-size:11px;background:var(--al);color:var(--a);border:1px solid var(--ab);border-radius:10px;padding:2px 8px">'+pp+' +'+radii[i]+'公里</span>';}).join('')+'</div>':'';
    const tagsHtml=s30.map(function(d){
      const safe=d.cp.replace(/\\/g,'\\\\').replace(/'/g,"\\'");
      return '<div class="ptag" onclick="ct(\''+safe+'\')">'+d.cp+(d.z?'<span class="ptag-c">'+d.z+'</span>':'')+'</div>';
    }).join('');
    const perfHtml=calcPkgPerf(p,getAdType());
    html+='<div class="pkg" data-tags="'+tagIds.replace(/"/g,'&quot;')+'" style="margin-bottom:12px">';
    html+='<div class="gr-pkg-title"><div class="pbadge">'+labels[pkg]+'</div><span style="font-size:12px;font-weight:600;color:var(--tx)">'+strat.name+'</span><button class="cp-all" onclick="cpAll(this)">複製全部 ('+s30.length+')</button></div>';
    html+=locHtml;
    html+='<div style="font-size:11px;color:var(--sub);margin-bottom:8px;padding:6px 10px;background:var(--s2);border-radius:6px;line-height:1.55">'+strat.desc+'</div>';
    html+='<div class="ptags">'+tagsHtml+'</div>';
    html+=perfHtml;
    html+='</div>';
  }
  out.innerHTML=html;
  setTimeout(function(){out.scrollIntoView({behavior:'smooth',block:'start'});},50);
}
(function(){
  const LEJU_DATA={
    '台北市':[{dist:'大安',price:132},{dist:'南港',price:112},{dist:'信義',price:111},{dist:'松山',price:107},{dist:'大同',price:102},{dist:'中山',price:101},{dist:'士林',price:93},{dist:'文山',price:89},{dist:'萬華',price:85}],
    '新北市':[{dist:'永和',price:77},{dist:'板橋',price:73},{dist:'中和',price:69},{dist:'新店',price:68},{dist:'新莊',price:62},{dist:'林口',price:56},{dist:'淡水',price:35}],
    '桃園市':[{dist:'桃園',price:43}],
    '新竹市':[{dist:'新竹',price:47}],
    '台中市':[{dist:'台中',price:41}]
  };
  window.showLejuPanel=function(city,highlightDist){
    const anchor=document.getElementById('leju-panel-anchor');
    if(!anchor)return;
    const old=document.getElementById('leju-panel');if(old)old.remove();
    const rows=LEJU_DATA[city];
    if(!rows||!rows.length)return;
    const sorted=[...rows].sort(function(a,b){return b.price-a.price;});
    const maxPrice=sorted[0].price;
    const normHL=(highlightDist||'').replace(/[區市鄉鎮]$/,'');
    function isCur(d){return d.replace(/[區市鄉鎮]$/,'')=== normHL||d===highlightDist;}
    const rowsHtml=sorted.map(function(r){
      const cur=isCur(r.dist);
      const pct=Math.round(r.price/maxPrice*100);
      return '<div class="leju-row"><span class="leju-label'+(cur?' is-current':'')+'">'+r.dist+'<\/span><div class="leju-bar-track"><div class="leju-bar-fill'+(cur?' is-current':'')+'" style="width:'+pct+'%"><\/div><\/div><span class="leju-value'+(cur?' is-current':'')+'">'+r.price+'<\/span><\/div>';
    }).join('');
    const panel=document.createElement('div');
    panel.id='leju-panel';
    panel.style.cssText='padding:10px 11px 9px;background:var(--s2);border:1px solid var(--bd);border-radius:var(--radius-card);margin-bottom:12px';
    panel.innerHTML='<div class="leju-header"><span class="leju-title">地區行情參考<\/span><span class="leju-unit">萬／坪　近一年均價<\/span><\/div><div class="leju-rows">'+rowsHtml+'<\/div><div class="leju-source">資料來源：樂居 leju.com.tw<\/div>';
    anchor.appendChild(panel);
  };
})();
function getTagMultiplier(cat,unitPrice,areaType){
  const TIERS=[
    {min:130,max:Infinity,w:{__rich:3.0,__re:1.5,__tech:1.2,__civil:0.5,__family:0.8,__general:0.6}},
    {min:90,max:129,w:{__rich:2.2,__re:1.8,__tech:1.5,__civil:0.8,__family:1.0,__general:0.8}},
    {min:65,max:89,w:{__rich:1.4,__re:2.0,__tech:1.8,__civil:1.2,__family:1.5,__general:1.0}},
    {min:50,max:64,w:{__rich:0.8,__re:2.0,__tech:1.6,__civil:1.8,__family:2.0,__general:1.2}},
    {min:0,max:49,w:{__rich:0.5,__re:1.8,__tech:1.2,__civil:2.0,__family:2.5,__general:1.5}}
  ];
  const AREA={
    luxury:{__rich:1.3,__re:1.0,__tech:0.9,__civil:0.6,__family:0.8,__general:0.7},
    tech:{__rich:0.8,__re:1.1,__tech:1.6,__civil:0.9,__family:1.2,__general:1.0},
    resort:{__rich:0.9,__re:1.0,__tech:0.8,__civil:1.1,__family:1.3,__general:1.4},
    general:{__rich:1.0,__re:1.0,__tech:1.0,__civil:1.0,__family:1.0,__general:1.0}
  };
  const VALID=['__rich','__re','__tech','__civil','__family','__general'];
  if(VALID.indexOf(cat)<0)return 1.0;
  const ar=AREA[areaType]||AREA.general;
  const up=(typeof unitPrice==='number'&&unitPrice>0)?unitPrice:55;
  const tier=TIERS.find(function(t){return up>=t.min&&up<=t.max;});
  const base=tier?tier.w[cat]:1.0;
  const raw=base*(ar[cat]||1.0);
  return Math.round(Math.min(3.5,Math.max(0.3,raw))*100)/100;
}
var EXTRA_TAGS=[
  {cat:'財務投資',en:'Finance / Investment',tags:[
    {en:'Investment',zh:'投資'},{en:'Financial planning',zh:'財務規劃'},
    {en:'Stock market',zh:'股票市場'},{en:'Mutual fund',zh:'共同基金'},
    {en:'Retirement planning',zh:'退休規劃'},{en:'Wealth management',zh:'財富管理'},
    {en:'Personal finance',zh:'個人理財'},{en:'Asset management',zh:'資產管理'},
    {en:'Passive income',zh:'被動收入'},{en:'Portfolio management',zh:'投資組合管理'}
  ]},
  {cat:'房產直接',en:'Real Estate Direct',tags:[
    {en:'Real estate',zh:'房地產'},{en:'Mortgage loan',zh:'房屋貸款'},
    {en:'Home buying',zh:'購屋'},{en:'Property investment',zh:'不動產投資'},
    {en:'New construction',zh:'新成屋'},{en:'Real estate broker',zh:'房仲'},
    {en:'House hunting',zh:'找房'},{en:'Home ownership',zh:'置產'},
    {en:'Real estate developer',zh:'建商'},{en:'Interest rate',zh:'利率'}
  ]},
  {cat:'親子家庭',en:'Parenting & Family',tags:[
    {en:'Parenting',zh:'親子教養'},{en:'New parents',zh:'新手父母'},
    {en:'Infant',zh:'嬰幼兒'},{en:'Baby shower',zh:'迎嬰派對'},
    {en:'Child education',zh:'兒童教育'},{en:'Educational toys',zh:'益智玩具'},
    {en:'Kindergarten',zh:'幼稚園'},{en:'Primary school',zh:'小學'},
    {en:'Child development',zh:'兒童發展'},{en:'Baby products',zh:'嬰兒用品'}
  ]},
  {cat:'健康運動',en:'Health & Fitness',tags:[
    {en:'Health & fitness',zh:'健康健身'},{en:'Running',zh:'跑步'},
    {en:'Gym',zh:'健身房'},{en:'Yoga',zh:'瑜伽'},
    {en:'Cycling',zh:'騎單車'},{en:'Swimming',zh:'游泳'},
    {en:'Healthy eating',zh:'健康飲食'},{en:'Wellness',zh:'健康生活'},
    {en:'Organic food',zh:'有機食品'},{en:'Sports nutrition',zh:'運動營養'}
  ]},
  {cat:'科技職場',en:'Tech & Career',tags:[
    {en:'Software engineering',zh:'軟體工程'},{en:'Information technology',zh:'資訊科技'},
    {en:'Artificial intelligence',zh:'人工智慧'},{en:'Cloud computing',zh:'雲端運算'},
    {en:'Data science',zh:'資料科學'},{en:'Semiconductor',zh:'半導體'},
    {en:'Electronics engineering',zh:'電子工程'},{en:'Startup',zh:'新創公司'},
    {en:'Tech industry',zh:'科技業'},{en:'Science park',zh:'科學園區'}
  ]},
  {cat:'汽車交通',en:'Automotive',tags:[
    {en:'Automobile',zh:'汽車'},{en:'Car dealership',zh:'汽車經銷商'},
    {en:'Electric vehicle',zh:'電動車'},{en:'Luxury vehicle',zh:'豪華車'},
    {en:'Tesla',zh:'特斯拉'},{en:'BMW',zh:'寶馬'},
    {en:'Mercedes-Benz',zh:'賓士'},{en:'Toyota',zh:'豐田'},
    {en:'Hyundai',zh:'現代'},{en:'Car insurance',zh:'車險'}
  ]},
  {cat:'旅遊度假',en:'Travel & Leisure',tags:[
    {en:'Travel',zh:'旅遊'},{en:'International travel',zh:'國際旅行'},
    {en:'Hotel',zh:'飯店'},{en:'Resort',zh:'度假村'},
    {en:'Vacation',zh:'假期'},{en:'Airbnb',zh:'民宿'},
    {en:'Tourism',zh:'觀光'},{en:'Cruise',zh:'郵輪'},
    {en:'Backpacking',zh:'背包旅行'},{en:'Sightseeing',zh:'觀光景點'}
  ]},
  {cat:'美食生活',en:'Food & Dining',tags:[
    {en:'Restaurants',zh:'餐廳'},{en:'Fine dining',zh:'精緻餐飲'},
    {en:'Food delivery',zh:'外送服務'},{en:'Cooking',zh:'烹飪'},
    {en:'Gourmet food',zh:'美食'},{en:'Baking',zh:'烘焙'},
    {en:'Coffee',zh:'咖啡'},{en:'Wine',zh:'葡萄酒'},
    {en:'Foodie',zh:'美食愛好者'},{en:'Brunch',zh:'早午餐'}
  ]},
  {cat:'精品奢華',en:'Luxury & Fashion',tags:[
    {en:'Luxury goods',zh:'精品'},{en:'Designer handbags',zh:'名牌包'},
    {en:'Jewelry',zh:'珠寶'},{en:'Chanel',zh:'香奈兒'},
    {en:'Louis Vuitton',zh:'路易威登'},{en:'Hermès',zh:'愛馬仕'},
    {en:'Fashion',zh:'時尚'},{en:'Luxury watches',zh:'名錶'},
    {en:'High-end shopping',zh:'高端購物'},{en:'Premium lifestyle',zh:'頂級生活'}
  ]},
  {cat:'職場創業',en:'Business & Entrepreneurship',tags:[
    {en:'Entrepreneurship',zh:'創業'},{en:'Small business',zh:'小型企業'},
    {en:'Marketing',zh:'行銷'},{en:'Leadership',zh:'領導力'},
    {en:'Career development',zh:'職涯發展'},{en:'Business networking',zh:'商業社交'},
    {en:'MBA',zh:'企業管理碩士'},{en:'Executive',zh:'高階主管'},
    {en:'Management consulting',zh:'管理顧問'},{en:'Real estate agent',zh:'房仲業務'}
  ]}
];
var _xShown={};
var XTAG_SHOW=5;
var _dbTagSet=null;
function _getDbTagSet(){if(_dbTagSet)return _dbTagSet;_dbTagSet=new Set(D.map(function(d){return d.cp;}));return _dbTagSet;}
function _xPickRandom(gi){
  var g=EXTRA_TAGS[gi];if(!g)return[];
  var dbSet=_getDbTagSet();
  var arr=g.tags.filter(function(t){return!dbSet.has(t.en);}).slice();
  for(var i=arr.length-1;i>0;i--){var j=Math.floor(Math.random()*(i+1));var tmp=arr[i];arr[i]=arr[j];arr[j]=tmp;}
  return arr.slice(0,XTAG_SHOW);
}
function _buildXtagChips(gi){
  return (_xShown[gi]||[]).map(function(t){
    var en=t.en.replace(/'/g,'&#39;');
    return '<span class="xtag-chip" onclick="copyXtag(this,\''+en+'\')">'+
      '<span class="xtag-chip-en">'+t.en+'<\/span>'+
      '<span class="xtag-chip-zh">'+t.zh+'<\/span><\/span>';
  }).join('');
}
function shuffleXtags(gi){
  _xShown[gi]=_xPickRandom(gi);
  var chips=document.getElementById('xtag-chips-'+gi);
  var btn=document.getElementById('xtag-shuf-'+gi);
  if(chips)chips.innerHTML=_buildXtagChips(gi);
  if(btn){btn.style.transform='rotate(180deg)';setTimeout(function(){btn.style.transform='';},280);}
}
function shuffleAllXtags(){
  EXTRA_TAGS.forEach(function(g,gi){shuffleXtags(gi);});
}
function renderExtraTags(){
  var sec=document.getElementById('extra-tags-section');
  if(!sec)return;
  EXTRA_TAGS.forEach(function(g,gi){_xShown[gi]=_xPickRandom(gi);});
  var html='<div class="xtag-section">';
  html+='<div class="xtag-hd"><span class="xtag-hd-title">延伸興趣（未收錄標籤）　共 '+EXTRA_TAGS.length+' 類　點擊複製<\/span>';
  html+='<button class="xtag-shuffle-all" onclick="shuffleAllXtags()">🔄 全部刷新<\/button><\/div>';
  html+='<div class="xtag-grid">';
  EXTRA_TAGS.forEach(function(g,gi){
    html+='<div class="xtag-card">';
    html+='<div class="xtag-card-hd"><div><div class="xtag-card-cat">'+g.cat+'<\/div><div class="xtag-card-en">'+g.en+'<\/div><\/div>';
    html+='<button class="xtag-card-shuf" id="xtag-shuf-'+gi+'" onclick="shuffleXtags('+gi+')" title="換一批">🔄<\/button><\/div>';
    html+='<div class="xtag-chips" id="xtag-chips-'+gi+'">'+_buildXtagChips(gi)+'<\/div>';
    html+='<\/div>';
  });
  html+='<\/div><\/div>';
  sec.innerHTML=html;
}
function copyXtag(el,text){
  if(navigator.clipboard){
    navigator.clipboard.writeText(text).then(function(){
      el.classList.add('copied');
      setTimeout(function(){el.classList.remove('copied');},900);
    });
  }
}
var _bldgAreaMap={'大安':'luxury','信義':'luxury','中山':'luxury','南港':'tech',
  '竹北':'tech','新竹市':'tech','淡水':'resort'};
function clearBldg(){
  var addrBlock=document.getElementById('addr-search-block');
  if(addrBlock){addrBlock.style.opacity='';addrBlock.style.pointerEvents='';}
  var an=document.getElementById('bldg-analysis-anchor');if(an)an.innerHTML='';
  var rp=document.getElementById('real-price-anchor');if(rp)rp.innerHTML='';
}
function showBldgAnalysis(item){
  var anchor=document.getElementById('bldg-analysis-anchor');
  if(!anchor||!item||!item.price)return;
  var distKey=item.distKey||'';
  var prof=getBuyerProfile(item.price,distKey,window.mapAreaType||'general');
  var card=document.createElement('div');
  card.className='lbp-card';
  card.style.marginTop='10px';
  var h='<div class="lbp-header"><span class="lbp-title">建案分析<\/span><\/div>';
  h+='<table style="width:100%;font-size:11px;border-collapse:collapse;line-height:1.7">';
  h+='<tr><td style="color:var(--sub);width:68px;vertical-align:top">建案名稱<\/td><td style="font-weight:600;color:var(--tx)">'+item.name+'<\/td><\/tr>';
  h+='<tr><td style="color:var(--sub);vertical-align:top">地點<\/td><td style="color:var(--tx)">'+(item.dist||'')+(item.city?' ・ '+item.city:'')+'<\/td><\/tr>';
  h+='<tr><td style="color:var(--sub);vertical-align:top">均價<\/td><td style="font-weight:700;color:var(--a-dark)">'+item.price+' 萬\/坪<\/td><\/tr>';
  if(prof.relativeNote){
    h+='<tr><td style="color:var(--sub);vertical-align:top">地區比對<\/td><td style="color:var(--tx)">'+prof.relativeNote+'<\/td><\/tr>';
  }
  h+='<tr><td style="color:var(--sub);vertical-align:top">案件定位<\/td><td style="color:var(--tx)"><b>'+prof.tierLabel+'<\/b><\/td><\/tr>';
  h+='<tr><td style="color:var(--sub);vertical-align:top">主力客群<\/td><td style="color:var(--tx)">'+prof.buyers.join('、')+'<\/td><\/tr>';
  h+='<tr><td style="color:var(--sub);vertical-align:top">推薦標籤<\/td><td style="color:var(--tx)">'+prof.catFocus+'<\/td><\/tr>';
  h+='<\/table><div class="lbp-source">分析依據：樂居成交均價 + 地區比對<\/div>';
  card.innerHTML=h;
  anchor.innerHTML='';
  anchor.appendChild(card);
}
function computeUserUnitPrice(){
  var ptype=document.querySelector('input[name="ptype"]:checked');
  if(!ptype)return null;
  var t=ptype.value;
  var s1=parseFloat(document.getElementById('sz1').value)||0;
  var s2=parseFloat(document.getElementById('sz2').value)||s1;
  var avgSize=(s1+s2)/2;
  if(t==='unit'){
    var u=parseFloat(document.getElementById('pu1').value)||0;
    return u>0?u:null;
  }else{
    var p1=parseFloat(document.getElementById('pt1').value)||0;
    var p2=parseFloat(document.getElementById('pt2').value)||p1;
    var totalAvg=(p1+p2)/2;
    return (totalAvg>0&&avgSize>0)?totalAvg/avgSize:null;
  }
}
(function(){
  var BLDG={
    '林口':[
      {n:'原森TWIN TOWERS',p:77.6},{n:'家悅美',p:67.9},{n:'長虹天聚',p:65.0},
      {n:'太子苑',p:62.5},{n:'立陽頤興',p:57.9},{n:'玄泰ONE',p:52.0},
      {n:'夢想之都',p:48.6},{n:'築禾琢立',p:45.0},{n:'杜拜The Palm Island',p:41.2},{n:'世界首席',p:38.9}
    ],
    '新莊':[
      {n:'富都新玉',p:76.1},{n:'勝旺新',p:74.8},{n:'泰鼎自由之丘',p:72.5},
      {n:'青青雙橡園',p:71.3},{n:'豐隆昕誠',p:70.8},{n:'宏璟榮華',p:70.4},
      {n:'永陞聯合公園',p:68.2},{n:'國泰新莊園',p:65.4},{n:'宏福新莊',p:51.0}
    ],
    '板橋':[
      {n:'三輝白昀',p:120.8},{n:'嘉廷家傳',p:84.0},{n:'君麟',p:79.6},
      {n:'新美齊画世代',p:76.7},{n:'板橋璞園的家',p:76.5},{n:'民生新埔',p:76.1},
      {n:'板橋大廈',p:66.4},{n:'板橋世家',p:61.9},{n:'鎮板橋',p:49.5},{n:'大庭新城',p:49.2}
    ],
    '大安':[
      {n:'台北之星',p:215.0},{n:'吾雙',p:201.6},{n:'大陸耑岫',p:180.0},
      {n:'大安A+',p:164.2},{n:'大安傳家',p:147.7},{n:'大安薈館',p:132.4},
      {n:'大安國宅',p:96.5},{n:'大安捷運廣場',p:95.9},{n:'大安1景',p:89.6}
    ],
    '信義':[
      {n:'法意大樓',p:199.1},{n:'信義之星',p:168.9},{n:'璞園信義',p:160.7},
      {n:'吉祥如藝',p:139.9},{n:'信義香榭',p:110.5},{n:'鼎真麗澤',p:91.0},{n:'信義國際',p:84.2}
    ],
    '中山':[
      {n:'中山Q1',p:114.1},{n:'帝璽',p:113.2},{n:'中山居',p:107.1},
      {n:'品中山',p:104.8},{n:'中山隱',p:104.1},{n:'中山官邸',p:100.6},
      {n:'太子中山大樓',p:86.0},{n:'中山民權大樓',p:73.4},{n:'中山首府',p:65.7}
    ],
    '南港':[
      {n:'世界明珠',p:146.7},{n:'國泰SKY PARK',p:139.9},{n:'國協天玥',p:126.8},
      {n:'東方大境',p:121.2},{n:'擎天森林',p:109.1},{n:'昆陽園',p:108.5},
      {n:'蘋果廣場',p:94.0},{n:'東方大地',p:81.7},{n:'南港軟體園區2',p:73.3},{n:'南港花園社區',p:63.1}
    ],
    '桃園':[
      {n:'合眾MRT',p:52.0},{n:'百俊吾双',p:50.9},{n:'金龍皇璽',p:50.9},
      {n:'益騏新画峰',p:49.2},{n:'匯聚',p:49.1},{n:'昭揚縱橫',p:41.3},
      {n:'長榮富豪',p:28.3},{n:'青天花園',p:28.2}
    ],
    '新竹市':[
      {n:'新東京',p:75.1},{n:'沅朔曦月',p:59.3},{n:'昌益文清',p:57.7},
      {n:'科技之心',p:53.7},{n:'米樂第',p:52.9},{n:'鴻築MM21',p:46.1},
      {n:'綠光金澤',p:49.1},{n:'和風大樓',p:43.6},{n:'雲冠天下',p:30.3}
    ],
    '竹北':[
      {n:'昌益御邸',p:91.9},{n:'豐采520',p:88.1},{n:'坤山臻真',p:74.0},
      {n:'若山若餘山',p:69.1},{n:'高峰匯',p:62.8},{n:'禾子旺星光',p:61.9},
      {n:'竹北101',p:43.2},{n:'金洋微上',p:39.3},{n:'六家高鐵B區',p:30.7},{n:'竹北雄觀',p:21.9}
    ],
    '永和':[
      {n:'頂溪大苑',p:109.8},{n:'M PLUS',p:99.9},{n:'漢皇River Sky',p:91.6},
      {n:'詠和心',p:80.1},{n:'黃金時代',p:64.0},{n:'恆豐大樓',p:63.6},
      {n:'鉅陞永麗花園',p:61.5},{n:'漢寶臺大苑',p:55.6},{n:'永和路華廈',p:49.7},{n:'永和成功',p:40.0}
    ],
    '中和':[
      {n:'漢皇方圓',p:114.7},{n:'和御',p:80.1},{n:'大同新紀元',p:78.4},
      {n:'玖原君邸',p:78.3},{n:'冠德心禾匯',p:78.0},{n:'遠雄CASA',p:70.3},
      {n:'永邦恆美',p:57.3},{n:'中和路公寓',p:49.0},{n:'中和100',p:44.2}
    ],
    '新店':[
      {n:'寶安JR',p:96.1},{n:'江陵天喆',p:94.8},{n:'華固譽誠',p:94.4},
      {n:'波爾多',p:69.7},{n:'敦南之森',p:65.5},{n:'新店臺北人',p:64.9},
      {n:'現代君址',p:62.9},{n:'家瑞敦南花園',p:48.4}
    ],
    '淡水':[
      {n:'十得向陽',p:41.6},{n:'翰林苑上水',p:39.5},{n:'新海城',p:38.6},
      {n:'台北灣銀河',p:37.4},{n:'天藝',p:37.2},{n:'和光五明',p:35.5},
      {n:'上水',p:34.4},{n:'伊東市',p:34.1},{n:'台北灣四季之旅',p:29.6},{n:'淡水人',p:26.0}
    ]
  };
  var DIST_CENTER={
    '林口':{lat:25.079,lng:121.382,city:'新北市'},
    '新莊':{lat:25.041,lng:121.459,city:'新北市'},
    '板橋':{lat:25.012,lng:121.466,city:'新北市'},
    '大安':{lat:25.026,lng:121.543,city:'台北市'},
    '信義':{lat:25.033,lng:121.564,city:'台北市'},
    '中山':{lat:25.064,lng:121.524,city:'台北市'},
    '南港':{lat:25.054,lng:121.607,city:'台北市'},
    '桃園':{lat:24.993,lng:121.301,city:'桃園市'},
    '新竹市':{lat:24.803,lng:120.968,city:'新竹市'},
    '竹北':{lat:24.846,lng:121.014,city:'新竹縣'},
    '永和':{lat:25.013,lng:121.519,city:'新北市'},
    '中和':{lat:25.003,lng:121.500,city:'新北市'},
    '新店':{lat:24.972,lng:121.541,city:'新北市'},
    '淡水':{lat:25.168,lng:121.448,city:'新北市'}
  };
  window._BLDG=BLDG;
  window._DIST_CENTER=DIST_CENTER;
  var _bldgPins=[];
  var _bldgMode='chart';
  var _curDist=null;
  var _curPrice=null;
  function _clearPins(){
    _bldgPins.forEach(function(m){try{m.remove();}catch(e){}});
    _bldgPins=[];
  }
  function _pinPos(distKey,idx,total){
    var c=DIST_CENTER[distKey];
    if(!c)return null;
    var angle=(idx/Math.max(total,1))*Math.PI*2-Math.PI/2;
    var r=0.003+(idx%3)*0.002;
    return {lat:c.lat+r*Math.sin(angle),lng:c.lng+r*Math.cos(angle)};
  }
  function _renderChart(wrap,distKey,userPrice){
    var rows=BLDG[distKey];
    if(!rows||!rows.length)return;
    var sorted=rows.slice().sort(function(a,b){return b.p-a.p;});
    var maxP=sorted[0].p;
    var nearestIdx=-1,nearestDiff=Infinity;
    if(userPrice!=null){
      sorted.forEach(function(r,i){var d=Math.abs(r.p-userPrice);if(d<nearestDiff){nearestDiff=d;nearestIdx=i;}});
    }
    var refLabel=userPrice!=null?'<span class="lbp-refline-label">你的案件 '+Math.round(userPrice)+'萬\/坪<\/span>':'';
    var card=document.createElement('div');
    card.className='lbp-card';
    card.innerHTML='<div class="lbp-header"><span class="lbp-title">'+distKey+'　建案價格分布<\/span><span class="lbp-subtitle">萬\/坪・近一年成交<\/span>'+refLabel+'<button class="lbp-mode-btn" onclick="window._bldgToggle()">🗺 地圖<\/button><\/div><div class="lbp-chart"><\/div><div class="lbp-source">資料來源：樂居 leju.com.tw<\/div>';
    wrap.appendChild(card);
    var chart=card.querySelector('.lbp-chart');
    sorted.forEach(function(r,i){
      var row=document.createElement('div');
      row.className='lbp-row'+(i===nearestIdx?' lbp-row--nearest':'');
      var pct=Math.round(r.p/maxP*100);
      row.innerHTML='<div class="lbp-name">'+r.n+'<\/div><div class="lbp-track"><div class="lbp-bar" style="width:'+pct+'%"><\/div><\/div><div class="lbp-price">'+r.p+'萬<\/div>';
      chart.appendChild(row);
    });
  }
  function _renderMap(wrap,distKey,userPrice){
    _clearPins();
    var rows=BLDG[distKey];
    if(!rows||!rows.length)return;
    var sorted=rows.slice().sort(function(a,b){return b.p-a.p;});
    var maxP=sorted[0].p,minP=sorted[sorted.length-1].p;
    var card=document.createElement('div');
    card.className='lbp-card';
    card.innerHTML='<div class="lbp-header"><span class="lbp-title">'+distKey+'　建案價格地圖<\/span><span class="lbp-subtitle">萬\/坪・近一年成交<\/span><button class="lbp-mode-btn" onclick="window._bldgToggle()">📊 圖表<\/button><\/div><div class="lbp-source" style="margin-top:2px">點選價格標籤查看建案名稱　資料：樂居<\/div>';
    wrap.appendChild(card);
    var m=window._map;
    if(!m)return;
    var ctr=DIST_CENTER[distKey];
    if(ctr)m.setView([ctr.lat,ctr.lng],14);
    var nearestPrice=null;
    if(userPrice!=null){
      nearestPrice=sorted.reduce(function(b,r){return Math.abs(r.p-userPrice)<Math.abs(b.p-userPrice)?r:b;},sorted[0]).p;
    }
    sorted.forEach(function(r,i){
      var pos=_pinPos(distKey,i,sorted.length);
      if(!pos)return;
      var ratio=maxP>minP?(r.p-minP)/(maxP-minP):0.5;
      var hue=Math.round((1-ratio)*210);
      var isN=(userPrice!=null&&r.p===nearestPrice);
      var outline=isN?'outline:2px solid #ff6b35;box-shadow:0 0 0 3px rgba(255,107,53,.3);':'';
      var html='<div style="padding:3px 8px;background:hsl('+hue+',65%,42%);color:#fff;border-radius:12px;font-size:10px;font-weight:700;white-space:nowrap;box-shadow:0 1px 4px rgba(0,0,0,.35);cursor:pointer;'+outline+'">'+r.p+'萬<\/div>';
      var icon=L.divIcon({className:'',html:html,iconAnchor:[28,12]});
      var mk=L.marker([pos.lat,pos.lng],{icon:icon}).addTo(m);
      mk.bindPopup('<b>'+r.n+'<\/b><br>均價 '+r.p+' 萬\/坪');
      _bldgPins.push(mk);
    });
  }
  window._bldgToggle=function(){
    _bldgMode=(_bldgMode==='chart'?'map':'chart');
    var wrap=document.getElementById('leju-bldg-panel');
    if(!wrap)return;
    wrap.innerHTML='';
    if(_bldgMode==='map'){_renderMap(wrap,_curDist,_curPrice);}
    else{_clearPins();_renderChart(wrap,_curDist,_curPrice);}
  };
  window.showLejuBldgPanel=function(distKey,userPrice){
    _curDist=distKey;_curPrice=userPrice;
    var wrap=document.getElementById('leju-bldg-panel');
    if(!wrap)return;
    wrap.innerHTML='';
    _clearPins();
    if(_bldgMode==='map'){_renderMap(wrap,distKey,userPrice);}
    else{_renderChart(wrap,distKey,userPrice);}
  };
})();
function getAdType(){
  const el=document.querySelector('input[name="adtype"]:checked');
  return el?el.value:'名單';
}
function getBuyerProfile(unitPrice,distKey,areaType){
  var bldg=window._BLDG&&window._BLDG[distKey];
  var distMedian=null;
  if(bldg&&bldg.length){
    var ps=bldg.map(function(r){return r.p;}).sort(function(a,b){return a-b;});
    distMedian=ps[Math.floor(ps.length/2)];
  }
  var tier,tierLabel,buyers,catFocus;
  if(unitPrice>=150){
    tier='ultra';tierLabel='豪宅頂奢';
    buyers=['高資產人士','企業主','私人銀行客戶','海外投資客'];
    catFocus='有錢興趣、頂奢生活、財富管理、私人銀行';
  }else if(unitPrice>=80){
    tier='high';tierLabel='高端';
    buyers=['高階主管','金融業高管','科技業高管','成功創業者'];
    catFocus='有錢興趣、理財投資、品質生活、換屋需求';
  }else if(unitPrice>=50){
    tier='midhigh';tierLabel='中高端換屋';
    buyers=['換屋族','科技業工程師','雙薪家庭','高階公務員'];
    catFocus='科技業興趣、換屋房地產、家庭教育、理財規劃';
  }else if(unitPrice>=30){
    tier='mid';tierLabel='中端首購';
    buyers=['首購族','公務員','年輕雙薪夫妻','有房需求家庭'];
    catFocus='首購需求、公務員族群、家庭生活、房地產直接';
  }else{
    tier='entry';tierLabel='平價首購';
    buyers=['社會新鮮人','租轉買族','剛性需求首購'];
    catFocus='年輕首購、生活機能、租轉買族';
  }
  if(areaType==='tech'&&(tier==='midhigh'||tier==='mid')){
    buyers=['竹科/桃科工程師','換屋族','雙薪工程師家庭','科技業上班族'];
    catFocus='科技業興趣（園區工程師）、換屋房地產、家庭生活、理財規劃';
  }
  if(areaType==='luxury'&&tier!=='mid'&&tier!=='entry'){
    catFocus='有錢興趣、頂奢生活、財富管理、豪宅購置';
  }
  if(areaType==='resort'){
    buyers=buyers.concat(['度假置產族','退休規劃者']);
    catFocus=catFocus+'、旅遊休閒、退休規劃';
  }
  var relativeNote='';
  if(distMedian&&unitPrice>0){
    var ratio=Math.round((unitPrice/distMedian-1)*100);
    if(ratio>15)relativeNote=distKey+'區均價約'+distMedian+'萬/坪，高出 '+ratio+'%（地區高端）';
    else if(ratio<-15)relativeNote=distKey+'區均價約'+distMedian+'萬/坪，低於 '+Math.abs(ratio)+'%（地區平價）';
    else relativeNote=distKey+'區均價約'+distMedian+'萬/坪，與周邊相近';
  }
  return {tier:tier,tierLabel:tierLabel,buyers:buyers,catFocus:catFocus,distMedian:distMedian,relativeNote:relativeNote};
}
function showRealPricePanel(distKey,city){
  var anchor=document.getElementById('real-price-anchor');
  if(!anchor)return;
  anchor.innerHTML='';
  var distName=(distKey||'')+'區';
  // Always render a link card first
  var linkCard=document.createElement('div');
  linkCard.className='lbp-card';linkCard.style.marginTop='10px';
  linkCard.innerHTML='<div class="lbp-header"><span class="lbp-title">實價登錄<\/span><span class="lbp-subtitle">'+distName+'<\/span><\/div>'+
    '<div style="font-size:11px;color:var(--sub);padding:6px 0 2px">直接查詢官方資料：<a href="https://plvr.land.moi.gov.tw/" target="_blank" style="color:var(--a);font-weight:600">內政部不動產成交案件查詢<\/a><\/div>'+
    '<div class="lbp-source">資料來源：內政部實價登錄<\/div>';
  anchor.appendChild(linkCard);
  // Attempt live data for New Taipei
  if(city==='新北市'){
    var url='https://data.ntpc.gov.tw/api/datasets/acce802d-58cc-4dff-9e7a-9ecc517f78be/json?page=0&size=10&filter='+encodeURIComponent("district eq '"+distName+"'");
    fetch(url).then(function(r){return r.json();}).then(function(data){
      if(!data||!data.length)return;
      var rows=data.filter(function(d){return parseInt(d['單價元\/平方公尺']||0)>0;}).slice(0,5);
      if(!rows.length)return;
      var card=document.createElement('div');
      card.className='lbp-card';card.style.marginTop='6px';
      var inner='<div class="lbp-header"><span class="lbp-title">近期實價登錄<\/span><span class="lbp-subtitle">'+distName+' 近期成交<\/span><\/div>';
      rows.forEach(function(d){
        var sqm=parseInt(d['單價元\/平方公尺'])||0;
        var ping=Math.round(sqm*3.3058/10000*10)/10;
        var ym=d['交易年月']||'';
        var yr=(parseInt(ym.substring(0,3))||0)+1911;
        var mo=parseInt(ym.substring(3,5))||0;
        var addr=(d['土地區段位置或建物區段門牌']||'').substring(0,18);
        inner+='<div style="display:grid;grid-template-columns:60px 1fr 48px;align-items:center;gap:6px;padding:5px 0;border-bottom:1px solid var(--bd);font-size:11px">';
        inner+='<span style="color:var(--sub)">'+yr+'/'+mo+'<\/span>';
        inner+='<span style="color:var(--tx);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+addr+'<\/span>';
        inner+='<span style="text-align:right;font-weight:600;color:var(--a-dark)">'+ping+'萬<\/span><\/div>';
      });
      inner+='<div class="lbp-source">資料來源：新北市政府開放資料 data.ntpc.gov.tw<\/div>';
      card.innerHTML=inner;
      anchor.appendChild(card);
    }).catch(function(){});
  }
}
function updateCriteria(){
  const box=document.getElementById('criteria-box');
  if(!box)return;
  const priceType=document.querySelector('input[name="ptype"]:checked').value;
  const s1=parseFloat(document.getElementById('sz1').value)||0;
  const s2=parseFloat(document.getElementById('sz2').value)||s1;
  const avgSize=(s1+s2)/2;
  let totalAvg=0,unitAvg=0;
  if(priceType==='total'){
    const p1=parseFloat(document.getElementById('pt1').value)||0;
    const p2=parseFloat(document.getElementById('pt2').value)||p1;
    totalAvg=(p1+p2)/2;
    if(avgSize>0)unitAvg=totalAvg/avgSize;
  }else{
    unitAvg=parseFloat(document.getElementById('pu1').value)||0;
    if(avgSize>0)totalAvg=unitAvg*avgSize;
  }
  const adType=getAdType();
  if(totalAvg===0&&unitAvg===0){box.style.display='none';return;}
  const priceLine=totalAvg>0?'總價約'+Math.round(totalAvg)+'萬'+(unitAvg>0?' / 單價約'+Math.round(unitAvg)+'萬/坪':''):unitAvg>0?'單價約'+Math.round(unitAvg)+'萬/坪':'';
  const sizeLine=(avgSize>0&&(totalAvg>0||unitAvg>0))?'坪數約'+Math.round(avgSize)+'坪':'';
  const at=window.mapAreaType||'general';
  const atLabel={tech:'科技園區',luxury:'精華豪宅',resort:'度假觀光',general:'一般住宅'}[at];
  const adDesc={'名單':'名單廣告 — 優先觸及有購屋意願用戶，以留資為轉換目標','訊息':'訊息廣告 — 導流 Messenger 對話，觸及高互動意願用戶','1510':'1510 影音 — 15秒短影片 + 靜態素材，擴大品牌觸及與案件印象'}[adType]||adType;
  const prof=getBuyerProfile(unitAvg||0,window.mapDistKey||'',at);
  const lines=['案件定位：'+prof.tierLabel+(unitAvg>0?' ('+Math.round(unitAvg)+'萬/坪)':'')+(atLabel?' ｜ '+atLabel+'地帶':'')];
  if(priceLine)lines.push('價格條件：'+priceLine);
  if(sizeLine)lines.push('坪數條件：'+sizeLine);
  if(prof.relativeNote)lines.push('地區比對：'+prof.relativeNote);
  lines.push('目標族群：'+prof.buyers.join('、'));
  lines.push('廣告形式：'+adDesc);
  lines.push('AI 優先選取：「'+prof.catFocus+'」，依地點×單價雙因子評分排序');
  box.innerHTML=lines.join('<br>');
  box.style.display='block';
  if(window.mapDistKey)showLejuBldgPanel(window.mapDistKey,computeUserUnitPrice());
}
function calcPkgPerf(p,adType){
  const at=window.mapAreaType||'general';
  // CPL benchmarks calibrated from real Taiwan real estate META ad data
  let cplLo,cplHi;
  if(p.isLuxury){cplLo=700;cplHi=1400;}
  else if(p.isHighEnd){cplLo=400;cplHi=800;}
  else if(p.isMid){cplLo=250;cplHi=550;}
  else{cplLo=180;cplHi=380;}
  if(at==='luxury'){cplLo=Math.round(cplLo*1.15);cplHi=Math.round(cplHi*1.15);}
  if(at==='tech'){cplLo=Math.round(cplLo*0.92);cplHi=Math.round(cplHi*0.92);}
  if(adType==='名單'){
    return '<div class="pkg-perf"><span class="perf-label">預計成效</span><span class="perf-item">預估 CPL：NT$'+cplLo+'～'+cplHi+'</span></div>';
  }else if(adType==='訊息'){
    return '<div class="pkg-perf"><span class="perf-label">預計成效</span><span class="perf-item">預估每次對話：NT$'+cplLo+'～'+cplHi+'</span></div>';
  }else{
    const cpmLo=p.isLuxury?120:p.isHighEnd?90:p.isMid?70:55;
    const cpmHi=p.isLuxury?200:p.isHighEnd?160:p.isMid?130:100;
    return '<div class="pkg-perf"><span class="perf-label">預計成效</span><span class="perf-item">預估 CPM：NT$'+cpmLo+'～'+cpmHi+'</span></div>';
  }
}
ry();rt();buildProjList();renderExtraTags();
'@

$htmlEnd = @'
</script></body></html>
'@

$fullHtml = $htmlPart1 + $jsFuncs + $htmlEnd
[System.IO.File]::WriteAllText("C:\Users\Authma\Desktop\Claude\META興趣標籤查詢工具.html", $fullHtml, $utf8NoBom)
$sz = (Get-Item "C:\Users\Authma\Desktop\Claude\META興趣標籤查詢工具.html").Length
Write-Host "Done! $([math]::Round($sz/1KB,0)) KB  Tags=$($tagList.Count)  YOU=$($youList.Count)  Rich=$nRich  Young=$nYoung"
