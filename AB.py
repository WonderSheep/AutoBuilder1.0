import sys
import re,os,time
import pandas as pd
import itertools
from playwright.sync_api import Playwright, sync_playwright, expect
from urllib.parse import urlparse, parse_qs

class IDCombinationSelector:#避投组合

    """
    任意数量N个ID的组合选择器（原12个ID版本的通用版）
    核心逻辑：初始化时一次性生成所有非空组合并缓存，后续查询直接取缓存，避免重复计算
    """
    def __init__(self, custom_ids):
        """
        初始化：传入N个ID，一次性生成所有组合并缓存
        :param custom_ids: N个ID组成的列表（N≥1，数字/字符串均可）
        """
        # 第一步：校验ID列表不能为空（替代原固定12个的校验）
        if not isinstance(custom_ids, list):
            raise TypeError(f"必须传入列表类型的ID集合！你当前传入的是{type(custom_ids)}")
        self.id_count = len(custom_ids)  # 动态记录ID的数量（N）
        if self.id_count == 0:
            raise ValueError("传入的ID列表不能为空！至少需要1个ID")
        self.custom_ids = custom_ids  # 保存原始N个ID
        self.total_valid = 2 ** self.id_count - 1  # 动态计算总有效组合数：2^N -1
        # 第二步：一次性生成所有组合并缓存（核心逻辑，仅执行一次）
        self.all_selected = self._generate_all_combinations()

    def _generate_all_combinations(self):
        """私有方法：生成所有非空组合（单ID→双ID→…→N个ID），仅初始化时调用"""
        all_comb = []
        # 动态循环：从1个ID到N个ID的升序组合（替代原固定1到13的循环）
        for select_num in range(1, self.id_count + 1):
            comb = itertools.combinations(self.custom_ids, select_num)
            all_comb.extend(list(comb))
        return all_comb

    def get_nth_choice(self, n):
        """
        公共查询方法：获取第n次的ID组合（直接从缓存中取，不重复生成）
        :param n: 目标查询次数（1≤n≤总组合数）
        :return: 第n次选用的ID列表
        """
        # 严格校验次数n的合法性
        if not isinstance(n, int):
            raise TypeError(f"查询次数必须是正整数！你输入的是{n}（类型：{type(n)}）")
        if n < 1 or n > self.total_valid:
            raise ValueError(
                f"查询次数超出有效范围！{self.id_count}个ID的有效次数是 1 ~ {self.total_valid}（当前输入：{n}）"
            )
        # 直接从缓存取值（索引n-1：列表从0开始，n从1开始）
        return list(self.all_selected[n - 1])

def get_chromium_path():
    if hasattr(sys, '_MEIPASS'):
        # 打包后：临时目录里的browser文件夹
        browser_root = os.path.join(sys._MEIPASS, "browser")
    else:
        # 未打包：脚本目录里的browser文件夹
        browser_root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "browser")
    
    # 动态找chromium文件夹（兼容带版本号的情况）
    chromium_folder = None
    for item in os.listdir(browser_root):
        item_path = os.path.join(browser_root, item)
        if os.path.isdir(item_path) and (item.startswith("chromium") or item == "chromium"):
            chromium_folder = item_path
            break
    
    if not chromium_folder:
        raise FileNotFoundError("❌ 未找到chromium文件夹，请检查browser目录！\n")
    
    # Windows下的chrome.exe路径
    chrome_exe = os.path.join(chromium_folder, "chrome.exe")
    # macOS替换成：os.path.join(chromium_folder, "Contents", "MacOS", "Chromium")
    # Linux替换成：os.path.join(chromium_folder, "chrome")
    
    if not os.path.exists(chrome_exe):
        raise FileNotFoundError(f"❌ 未找到chromium可执行文件：{chrome_exe}\n")
    
    return chrome_exe

def get_current_folder():
    # 打包后，sys._MEIPASS是exe内部临时目录，实际运行目录是exe所在目录
    if hasattr(sys, '_MEIPASS'):
        # 返回exe所在的目录（不是临时目录）
        return os.path.dirname(sys.executable)
    else:
        # 未打包时，返回脚本所在目录
        return os.path.dirname(os.path.abspath(__file__))

def read_excel_file():
    current_folder = get_current_folder()  # 用适配后的路径
    df = None  # 初始化df，避免未定义
    for filename in os.listdir(current_folder):
        file_full_path = os.path.join(current_folder, filename)
        is_valid_file = (
            os.path.isfile(file_full_path) 
            and filename.endswith(".xlsx") 
            and not filename.startswith(("~$", "$"))
        )
        
        if is_valid_file:
            try:
                df = pd.read_excel(file_full_path, engine="openpyxl", dtype=str)
                print(f"✅ 成功读取文件：{filename}\n")
                break  # 只读取第一个有效文件，保留你的break逻辑
            except Exception as e:
                print(f"❌ 读取文件 {filename} 失败：{e}\n")
                continue
        else:
            continue
    
    # 容错：没找到有效文件时提示并退出
    if df is None:
        print("❌ 未找到有效.xlsx文件，请确保exe同目录有非临时的xlsx文件！\n")
        sys.exit(1)  # 退出脚本，避免后续报错
    return df
             
def read_txt_file():
    """
    读取当前目录下第一个有效TXT文件（每行一个内容），直接返回清洗后的列表
    适配：每行一个内容的TXT，自动剔除空行/换行符/首尾空格
    返回：list - 每行内容对应的列表；无有效文件则退出脚本
    """
    current_folder = get_current_folder()
    # 初始化最终要返回的列表
    result_list = []
    
    for filename in os.listdir(current_folder):
        file_full_path = os.path.join(current_folder, filename)
        # 校验有效TXT文件：是文件、.txt结尾、非临时文件
        is_valid_file = (
            os.path.isfile(file_full_path) 
            and filename.endswith(".txt") 
            and not filename.startswith(("~$", "$"))
        )
        
        if is_valid_file:
            try:
                # 优先utf-8编码读取（适配大多数场景）
                with open(file_full_path, "r", encoding="utf-8") as f:
                    # 核心：逐行读取，清洗后存入列表
                    for line in f.readlines():
                        clean_line = line.strip()  # 剔除换行符/首尾空格
                        if clean_line:  # 跳过空行
                            result_list.append(clean_line)
                print(f"✅ 成功读取避投包人群文件：{filename}\n")
                print(f"✅ 共计 {len(result_list)} 个避投人群包，最多有 {2 ** len(result_list) - 1} 种避投组合\n")
                return result_list  # 直接返回列表，无需break
            
            # Windows中文TXT常见：utf-8失败则重试gbk编码
            except UnicodeDecodeError:
                try:
                    with open(file_full_path, "r", encoding="gbk") as f:
                        for line in f.readlines():
                            clean_line = line.strip()
                            if clean_line:
                                result_list.append(clean_line)
                    print(f"✅ 成功读取TXT文件（GBK编码）：{filename}")
                    print(f"✅ TXT共提取 {len(result_list)} 条有效内容\n")
                    return result_list
                
                except Exception as e:
                    print(f"❌ 读取TXT文件 {filename} 失败（编码错误）：{e}\n")
                    continue
            
            # 其他异常（文件损坏等）
            except Exception as e:
                print(f"❌ 读取TXT文件 {filename} 失败：{e}\n")
                continue
    
    # 容错：未找到有效TXT文件
    print("❌ 未找到有效.txt文件，请确保exe同目录有非临时的txt文件！\n")
    sys.exit(1)

def run_adq(playwright: Playwright,df,id_selector):

    def get_url_param(url, param_name):
        """
        从URL中提取指定键名的查询参数值
        :param url: 完整的URL字符串(比如page.url返回的内容)
        :param param_name: 要提取的参数键名(比如"project_id"、"name"、"type")
        :return: 指定键名的参数值(字符串,无该参数则返回None;有多个值返回第一个)
        """
        # 步骤1：解析URL，拆分出查询参数部分（?后面的内容）
        parsed_url = urlparse(url)
        # 步骤2：解析查询参数为字典（key: [value1, value2,...]，自动处理重复参数）
        query_params = parse_qs(parsed_url.query)
        # 步骤3：提取指定键名的参数值（无则返回None，有则取第一个值）
        param_value = query_params.get(param_name, [None])[0]
        return param_value

    ad_count_wx = 1 #微信搭建的拢共第几条广告，当前已有的WX广告数+1

    #ad_count_gdt = 1 #非微信搭建的拢共第几条广告，当前已有的GDT广告数+1
    
    logo = "KFC宅急送" if input("品牌形象若是 KFC宅急送 输入大写 Y；不是，直接 回车\n") == "Y"  else "肯德基"  #品牌形象

    #dp_LINK = "" #deeplink的组件名称

    offl_Lpage = "" #用于视频号-竖版视频的官方落地页

    user_input = input("行动按钮 若是 立即购买 直接 回车；若不是，请输入 并回车\n").strip()
    
    action_BTN = "立即购买" if user_input == "" else user_input  #行动按钮

    first_reply = input("朋友圈 的首评回复 若没有 直接 回车；若有，请输入 并回车\n").strip() #首评回复

    float_card = input("视频号-竖版视频 的浮层卡片 若没有 直接 回车；若有，请输入 并回车\n").strip() #用于竖版视频的浮层卡片名称，填写的是文案
    
    tv_tag = input("视频号-竖版视频 的标签 若没有 直接 回车；若有，请输入标签的其中一个 并回车\n").strip() #用于竖版视频的浮层卡片名称，填写的是文案

    print("Not ready yet\n")
    
    def WX_friends_card_bp() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="客服问答").click()
        if first_reply == "" :
            page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="首评回复").click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="行动按钮").click()
        page.locator("span.odc-text.ellipsis").filter(has_text="行动按钮").click()
        page.get_by_text("请选择按钮文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()
        if not first_reply == "" :
            page.locator("span.odc-text.ellipsis").filter(has_text="首评回复").click()
            page.get_by_text(first_reply).first.click()
    
    def WX_friends_card_vedio() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="客服问答").click()
        if first_reply == "" :
            page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="首评回复").click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="行动按钮").click()
        page.locator("span.odc-text.ellipsis").filter(has_text="行动按钮").click()
        page.get_by_text("请选择按钮文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()
        if not first_reply == "" :
            page.locator("span.odc-text.ellipsis").filter(has_text="首评回复").click()
            page.get_by_text(first_reply).first.click()

    def WX_friends_windows() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="图文链接").click()
        if first_reply == "" :
            page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="首评回复").click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="标签").click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="文字链").click()
        page.locator("span.odc-text.ellipsis").filter(has_text="文字链").click()
        page.get_by_text("请选择文字链文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()
        if not first_reply == "" :
            page.locator("span.odc-text.ellipsis").filter(has_text="首评回复").click()
            page.get_by_text(first_reply).first.click()

    def WX_friends_shuban_bp() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="图文链接").click()
        if first_reply == "" :
            page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="首评回复").click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="标签").click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="文字链").click()
        page.locator("span.odc-text.ellipsis").filter(has_text="文字链").click()
        page.get_by_text("请选择文字链文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()
        if not first_reply == "" :
            page.locator("span.odc-text.ellipsis").filter(has_text="首评回复").click()
            page.get_by_text(first_reply).first.click()
    
    def WX_Sub_bp() -> None:
        page.get_by_text("营销组件").first.click()
        #微信新闻插件
        #若选了 1注释 2不注释；
        #若没选 1不注释 2注释
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="行动按钮").click() # 1
        #page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="客服问答").click()# 2
        page.locator("span.odc-text.ellipsis").filter(has_text="行动按钮").click()
        page.get_by_text("请选择按钮文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()
    
    def WX_Sub_vedio() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="图文链接").click()#关闭图文链接
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="行动按钮").click()#打开行动按钮
        page.locator("span.odc-text.ellipsis").filter(has_text="行动按钮").click()
        page.get_by_text("请选择按钮文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()

    def WX_minipro_shuban_bp() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="图文链接").click()#关闭图文链接
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="标签").click()#关闭标签
        
    def WX_TV_shuban_vedio() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="图文链接").click()#关闭图文链接
        if tv_tag == "":
            page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="标签").click()#关闭标签
        page.locator("span.odc-text.ellipsis").filter(has_text="浮层卡片").click()
        page.locator("span.tw-text-xs.tw-text-text-secondary.tw-font-semibold.tw-truncate").filter(has_text=float_card).click()
        if not tv_tag == "":
            page.locator("span.odc-text.ellipsis").filter(has_text="标签").click()
            page.locator("div.tw-inline-flex.tw-items-center.tw-cursor-pointer.tw-transition-colors").filter(has_text=re.compile(rf"^{tv_tag}$")).first.click()
        #page.locator("span.odc-text.ellipsis").filter(has_text="标签").click()

    def GDT_shuban_bp() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="客服问答").click()#关闭客服问答
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="标签").click()#关闭标签
        page.locator("span.odc-text.ellipsis").filter(has_text="行动按钮").click()
        page.get_by_text("请选择按钮文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()

    def GDT_hengban_bp() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="客服问答").click()#关闭客服问答
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="标签").click()#关闭标签
        page.locator("span.odc-text.ellipsis").filter(has_text="行动按钮").click()
        page.get_by_text("请选择按钮文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()
    
    def GDT_flash_video() -> None:
        page.get_by_text("营销组件").first.click()
        page.locator("div.x-comp-overv-info span.odc-text").filter(has_text="客服问答").click()#关闭客服问答
        page.locator("span.odc-text.ellipsis").filter(has_text="行动按钮").click()
        page.get_by_text("请选择按钮文案").click()
        page.get_by_role("textbox", name="搜索").fill(action_BTN)
        page.wait_for_selector(f'div.selection-name[data-value="{action_BTN}"]',timeout=3000).click()
        page.get_by_role("button", name="确定").click()
   
    COMPONENT_MAP = {
        "朋友圈-卡片广告-横版大图-行动按钮" : WX_friends_card_bp,
        "朋友圈-卡片广告-横版大图" : WX_friends_card_bp,
        "朋友圈-卡片广告-横版视频-行动按钮" : WX_friends_card_vedio,
        "朋友圈-卡片广告-横版视频" : WX_friends_card_vedio,
        "朋友圈-竖版大图": WX_friends_shuban_bp,
        "朋友圈-橱窗广告-图片": WX_friends_windows,
        "订阅号消息列表-横版大图": WX_Sub_bp,
        "订阅号消息列表-横版视频": WX_Sub_vedio,
        "小程序封面广告" :WX_minipro_shuban_bp,
        "视频号-竖版视频" : WX_TV_shuban_vedio,
        "视频号评论区广告" : WX_TV_shuban_vedio,
        "竖版大图" : GDT_shuban_bp,
        "横版大图" : GDT_hengban_bp,
        "闪屏视频" : GDT_flash_video
    }

    PAGE_POSITION_MAP = {
    "朋友圈-卡片广告-横版大图-行动按钮":"卡片广告 横版大图 16:9",
    "朋友圈-卡片广告-横版大图":"卡片广告 横版大图 16:9",
    "朋友圈-卡片广告-横版视频-行动按钮":"卡片广告 横版视频 16:9",
    "朋友圈-卡片广告-横版视频":"卡片广告 横版视频 16:9",
    "朋友圈-竖版大图":"竖版大图 9:16",
    "朋友圈-橱窗广告-图片":"橱窗广告 - 图片",
    "订阅号消息列表-横版大图":"横版大图 16:9",
    "订阅号消息列表-横版视频":"横版视频 16:9",
    "小程序封面广告" :"竖版大图 9:16",
    "视频号-竖版视频" : "竖版视频 9:16",
    "视频号评论区广告" : "竖版视频 9:16",
    "竖版大图":"竖版大图 9:16",
    "横版大图":"横版大图 16:9",
    "闪屏视频":"闪屏视频 9:16"
    }

    #begin
    try:
        chrome_exe_path = get_chromium_path()
        print(f"✅ 找到Chromium路径：{chrome_exe_path}\n")
        # 启动浏览器
        browser = playwright.chromium.launch(
            executable_path=chrome_exe_path,
            headless=False  # 显示浏览器窗口，方便排查
        )
        # 创建上下文（保留你的原有逻辑）
        context = browser.new_context(storage_state="auth_state_adq.json") if os.path.exists("auth_state_adq.json") else browser.new_context()
    except Exception as e:
        print(f"❌ 启动浏览器失败：{e}\n")
        sys.exit(1)

    #region登录
    page = context.new_page()
    #"""
    page.goto(f"https://ad.qq.com")
    input("确保当前已处于登录状态后，按下回车开始搭建！\n")
    context.storage_state(path="auth_state_adq.json")#储存登录信息
    #"""
    #endregion

    for index,row in df.iterrows():

        #media表用
        #brand = row.iloc[3] #品牌
        strategy_ID = row.iloc[0] #策略ID
        campaign_NM = row.iloc[2] #活动名称
        media = row.iloc[13] #媒体
        page_PST = row.iloc[15] #点位
        audience = row.iloc[18] #脱敏人群
        creative_NM = row.iloc[20] #创意名称
        city = row.iloc[23] #区域
        #sell_Type = row.iloc[26] #购买类型
        mini_link = row.iloc[31] #小程序链接&掩码
        landing_type = row.iloc[32] #落地页类型
        #basic_LP = row.iloc[27] #打底落地页
        dp_LINK = row.iloc[29] #deeplink #目前用，直接在平台上选择dp
        #imp_TLink = row.iloc[35] #曝光监测链接
        #clk_Tlink = row.iloc[36] #点击监测链接
        
        #copy用
        account_id = row.iloc[40]# 广告账户
        copy_AD = row.iloc[42] #复制的广告

        #创意用
        asset_NM = row.iloc[44] #图片或视频名称
        copywriting = row.iloc[45] #文案

        #前端自定义人群
        #audience_MD = row.iloc[44] #媒体人群
        #audience_ID = row.iloc[46] #人群ID
        audience_tag = row.iloc[46] #人群标签

        print(f"创建账户:{account_id}")

        unit_NM = f"{strategy_ID}_{campaign_NM}_{media}_{page_PST}_{creative_NM}_{audience}_{audience_tag}_{city}"
        
        #print(unit_NM)
        page.goto(f"https://ad.qq.com/atlas/{account_id}/addelivery/adgroups-add?ref_adgroup_id={copy_AD}")
        
        time.sleep(3)
        #region广告
        
        #人群定向 排除人群
        
        if not id_selector is None :
            #try:
            #page.get_by_role("button", name="排除人群").first.dblclick(delay=350)
            page.locator('h3.title[title="排除人群"]').click()
            #except Exception :
                #page.get_by_role("button", name="全部定向").click()
                #page.get_by_role("button", name="排除人群").first.click()
            page.get_by_role("textbox", name="搜索用户群").click()
            #time.sleep(0.25)
            ad_count = ad_count_wx #if media == "微信" else ad_count_gdt
            avoid_list = id_selector.get_nth_choice(ad_count)
            for avoid in avoid_list :
                page.get_by_role("textbox", name="搜索用户群").fill(avoid)
                page.locator('tr[data-rowindex="0"] span.spaui-checkbox-indicator').first.click()
                time.sleep(0.15)
            ad_count_wx += 1
            
            """
            if media == "微信" : 
                print(f"微信的第 {ad_count_wx} 条广告")
                ad_count_wx += 1 
            else : 
                print(f"GDT的第 {ad_count_gdt} 条广告")
                ad_count_gdt += 1
            """

             #关闭高价值人群范围探索
            page.locator('span.spaui-switch-helper').nth(2).click() #有问题
        
        #人群定向 定向人群
        
        #选择监测链接组
        page.locator('div.spaui-selection-item-content.in').filter(has_text=re.compile(r"^请选择监测链接组$")).click()
        page.locator('span.name').filter(has_text="仅点击监测-DID").click()
        
        
        #项目名称
        #page.get_by_role("textbox", name="广告名称仅用于管理广告，不会对外展示").click()
        page.get_by_role("textbox", name="广告名称仅用于管理广告，不会对外展示").fill(unit_NM)

        #保存
        time.sleep(0.5)#等待0.5s确认填写完毕
        page.get_by_role("button", name="提交并新建创意").click()
        """
        cre_button = page.locator("button.spaui-button.spaui-button-primary").filter(has_text=re.compile(r"^创建创意$"))
        submit_btn=page.locator("button.spaui-button.spaui-button-primary.spaui-button-round").filter(has_text=re.compile(r"^确认并提交$"))

        if submit_btn.count() > 0:
            submit_btn.click()
            cre_button.click()
        else:
            cre_button.click()
        """
        submit_btn = page.locator("button", has_text=re.compile(r"确认并提交"))
        cre_button = page.locator("button", has_text=re.compile(r"^创建创意$"))
        
        if not id_selector is None :
            submit_btn.click()
            cre_button.click()
        else :
            cre_button.click()  
        
        #endregion

        #region创意

        #选点位
        page.locator("button#creative-type-btn").click()
        if page.locator("span.odc-text.ellipsis").filter(has_text=PAGE_POSITION_MAP[page_PST]).count() == 0:
            page.locator("span.spaui-switch-helper").click()
        page.locator("span.odc-text.ellipsis").filter(has_text=PAGE_POSITION_MAP[page_PST]).click()
        page.get_by_role("button", name="确定").click()

        #图片or视频
        if page_PST == "朋友圈-橱窗广告-图片":
            #page.locator("div.flex.x-anchor-item").first.click()
            page.locator("div.spaui-form-group-prefix button.x-filter-btn.spaui-button.spaui-button-text.spaui-button-sm.with-icon").first.click()
            page.get_by_role("textbox", name="请输入名称/ID").fill(asset_NM)#填图片或视频名称
            page.get_by_role("textbox", name="请输入名称/ID").press("Enter")
            time.sleep(0.5)
            page.locator('button[data-hottag="ComponentMediaSelector.changeImage"]').first.click(force=True)
            page.get_by_role("textbox", name="请输入小程序链接").click()
            page.get_by_role("textbox", name="请输入小程序链接").fill(mini_link)
            page.get_by_role("textbox", name="小程序链接", exact=True).click()
            page.get_by_role("textbox", name="小程序链接", exact=True).fill(mini_link)
            page.get_by_role("button", name="新建至账户").click()
        else :
            #page.locator("div.flex.x-anchor-item").first.click()
            page.locator("div.spaui-form-group-prefix button.x-filter-btn.spaui-button.spaui-button-text.spaui-button-sm.with-icon").first.click()
            page.get_by_role("textbox", name="请输入名称/ID").fill(asset_NM)#填图片ID
            page.get_by_role("textbox", name="请输入名称/ID").press("Enter")
            page.locator("div.odc-titlebar-desc.odc-line-clamp.break-all").filter(has_text=re.compile(rf"^ID:{asset_NM}$")).first.click(force=True)

        #文案
        if page.get_by_text("文案").count() > 0 :
            page.get_by_text("文案").click()
            page.locator("div.meta-input").click()
            page.locator("div.meta-input").fill(copywriting)#填文案

        #落地页
        if not page_PST == "朋友圈-橱窗广告-图片": #and not page_PST == "视频号-竖版视频" :
            page.locator("span").filter(has_text=re.compile(r"^落地页$")).click()
            if landing_type == "App" :
                page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^应用直达$")).click()
                page.locator("span.ellipsis.odc-text.odc-text-small.ellipsis.has-ellipsis").filter(has_text=re.compile(rf"^直达链接：{dp_LINK}$")).click(force=True)
                #page.locator("span.odc-tag-text.odc-text-mini").filter(has_text=re.compile(rf"^{dp_LINK}$")).click()
            else :
                page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^微信小程序$")).click()
                page.locator("button.x-elem-add-btn.spaui-button.spaui-button-default.with-icon").first.click()
                page.get_by_role("textbox", name="gh 开头的小程序原始 ID").fill("gh_f01f85672b87")#肯德基小程序
                page.get_by_role("textbox", name="请输入小程序链接").fill(mini_link)
                page.get_by_role("button", name="新建至账户").click()
        """
        elif  page_PST == "视频号-竖版视频" :    #视频号跳官方落地页已废弃？？
            page.locator("span").filter(has_text=re.compile(r"^落地页$")).click()
            #page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^微信小程序$")).click()
            #page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^应用直达$")).click()
            #page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^官方落地页$")).click()  官方落地页已废弃
            #page.locator("span.odc-text.odc-text-small.ellipsis").filter(has_text=re.compile(rf"^{offl_Lpage}$")).click()
            page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^微信小程序$")).click()
            page.locator("button.x-elem-add-btn.spaui-button.spaui-button-default.with-icon").first.click()
            #page.get_by_role("textbox", name="gh 开头的小程序原始 ID").click()
            page.get_by_role("textbox", name="gh 开头的小程序原始 ID").fill("gh_f01f85672b87")#肯德基小程序
            #page.get_by_role("textbox", name="请输入小程序链接").click()
            page.get_by_role("textbox", name="请输入小程序链接").fill(mini_link)
            page.get_by_role("button", name="新建至账户").click()
        """
        
        #品牌形象
        page.get_by_text("品牌形象", exact=True).click()
        if page_PST == "视频号-竖版视频" or page_PST == "视频号评论区广告":
            #page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^视频号$")).click()
            page.locator("div.h-full.flex-col.gap-4.px-12.justify-center.relative.odc-hover.odc-flex.odc-frame.flex.flex-row span.ellipsis.odc-text.odc-text-small.ellipsis").filter(has_text=re.compile(r"^肯德基$")).first.click()
        else :
            page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^自定义$")).click()
            page.locator("div.h-full.flex-col.gap-4.px-12.justify-center.relative.odc-hover.odc-flex.odc-frame.flex.flex-row span.ellipsis.odc-text.odc-text-small.ellipsis").filter(has_text=re.compile(rf"^{logo}$")).first.click()

        #营销组件
        COMPONENT_MAP[page_PST]()

        #创意名称
        page.get_by_text("创意设置").click()
        target_input = page.wait_for_selector("input.meta-input.spaui-input.has-normal",timeout=300000)
        target_input.click()
        target_input.press("ControlOrMeta+a")
        target_input.fill(unit_NM)

        #提交
        page.get_by_role("button", name="提交创意").click()
        page.get_by_role("button", name="返回创意管理").click()

        #endregion

        #去编辑
        page.get_by_role("button", name="编辑").first.wait_for(state="visible", timeout=300000)
        jump_url = page.get_by_role("button", name="编辑").first.get_attribute("href")
        adgroup_id = get_url_param(jump_url, "adgroup_id")
        dynamic_creative_id = get_url_param(jump_url, "dynamic_creative_id")
        #print(f"广告ID : {adgroup_id}")
        #print(f"创意ID : {dynamic_creative_id}")

        #填入ID
        row.iloc[42] = adgroup_id
        row.iloc[43] = dynamic_creative_id
        
        print(f"第{index+1}条广告 : {unit_NM} 创建成功\n")
    
    input("广告创建完成，plz press enter and continue")
    #输出新文件
    df = df.iloc[:, :44]
    df.to_excel("media_id_检查无误后可上传MOP.xlsx", index=False, engine="openpyxl")

    context.close()
    browser.close()

def run_adq_cre_template(playwright: Playwright,df,id_selector):

    #begin
    try:
        chrome_exe_path = get_chromium_path()
        print(f"✅ 找到Chromium路径：{chrome_exe_path}\n")
        # 启动浏览器
        browser = playwright.chromium.launch(
            executable_path=chrome_exe_path,
            headless=False  # 显示浏览器窗口，方便排查
        )
        # 创建上下文（保留你的原有逻辑）
        context = browser.new_context(storage_state="auth_state_adq.json") if os.path.exists("auth_state_adq.json") else browser.new_context()
    except Exception as e:
        print(f"❌ 启动浏览器失败：{e}\n")
        sys.exit(1)

    #region登录
    page = context.new_page()
    #"""
    page.goto(f"https://ad.qq.com")

    input("确保当前已处于登录状态后，按下回车开始创建定向模版！\n")

    context.storage_state(path="auth_state_adq.json")#储存登录信息

    account_id = df.iloc[0, 0] if not df.empty else df.columns[0]
    
    print(f"创建定向模版账户为：{account_id}")

    page.goto(f"https://ad.qq.com/atlas/{account_id}/addelivery/adgroups-add")
    #"""
    #endregion

    num = 2 ** id_selector.id_count - 1
    
    print(f"预计创建 {num} 个定向模版")
    
    page.get_by_role("button",name="展开更多选项").click()
    page.get_by_role("button",name="手动定向").click()
    page.get_by_role("button",name="使用手动定向").click()
    page.get_by_role("button", name="CPM", exact=True).click()
    #page.locator('button.spaui-button.spaui-button-default[data-value="4"]').click()

    for i in range(num):
        
        if i == 0 :
            page.get_by_role("button",name="全部定向").click()
            page.get_by_role("button",name="排除人群").click()
            
        if not i == 0 :
            page.locator("a.spaui-cursor-pointer").filter(has_text=re.compile(r"^清空$")).click()
        
        avoid_list = id_selector.get_nth_choice(i+1)
        
        for avoid in avoid_list :
            
            page.get_by_role("textbox", name="搜索用户群").fill(avoid)
            time.sleep(0.8)
            page.locator('tr[data-rowindex="0"] span.spaui-checkbox-indicator').first.click()
                
        if i == 0 :
            page.get_by_role("button",name="确定").click()
            page.locator('h3.title[title="排除人群"]').click()
        
        page.get_by_role("button",name="保存为定向模版").click()
                
        #template_NM =  page.locator('input.meta-input.spaui-input.has-normal[name][type="text"]')
        template_NM = page.get_by_role("textbox", name="请输入定向模版名称，最多50字")
        
        template_NM.fill(f"No{i+1}template{int(time.time())}")

        page.get_by_role("button", name="确定").click()

        print(f"第{i+1}条定向模版 : 创建成功\n")

    input("所有定向模版创建成功，press Enter and quit")

def run_adq_replace(playwright: Playwright,df):

    #begin
    try:
        chrome_exe_path = get_chromium_path()
        print(f"✅ 找到Chromium路径：{chrome_exe_path}\n")
        # 启动浏览器
        browser = playwright.chromium.launch(
            executable_path=chrome_exe_path,
            headless=False  # 显示浏览器窗口，方便排查
        )
        # 创建上下文（保留你的原有逻辑）
        context = browser.new_context(storage_state="auth_state_adq.json") if os.path.exists("auth_state_adq.json") else browser.new_context()
    except Exception as e:
        print(f"❌ 启动浏览器失败：{e}\n")
        sys.exit(1)

    #region登录
    page = context.new_page()
    #"""
    page.goto(f"https://ad.qq.com")
    input("确保当前已处于登录状态后，按下回车开始替换！\n")
    context.storage_state(path="auth_state_adq.json")#储存登录信息
    #"""
    #endregion

    for index,row in df.iterrows():

        #brand = row.iloc[3] #品牌
        account_ID = row.iloc[0] #账户ID
        ad_ID = row.iloc[1] #广告ID
        unit_ID = row.iloc[2] #创意ID
        asset_NM = row.iloc[3] if 3 < len(row) else "" #替换图片
        copywriting = row.iloc[4]  if 4 < len(row) else "" #文案
        landing = row.iloc[5]  if 5 < len(row) else "" #落地页
        action_btn = row.iloc[6]  if 6 < len(row) else "" #落地页
        

        page.goto(f"https://ad.qq.com/atlas/{account_ID}/delivery-page/creatives-update?adgroup_id={ad_ID}&dynamic_creative_id={unit_ID}")
        
        if pd.notna(asset_NM) and str(asset_NM).strip() != "" :
            page.get_by_role("button", name="清除").first.click()
            page.locator("div.spaui-form-group-prefix button.x-filter-btn.spaui-button.spaui-button-text.spaui-button-sm.with-icon").first.click()
            page.get_by_role("textbox", name="请输入名称/ID").fill(asset_NM)#填图片或视频名称
            page.get_by_role("textbox", name="请输入名称/ID").press("Enter")
            page.locator("div.odc-titlebar-desc.odc-line-clamp.break-all").filter(has_text=re.compile(rf"^ID:{asset_NM}$")).first.click(force=True)
        
        #文案
        if page.get_by_text("文案").count() > 0 and pd.notna(copywriting) and str(copywriting).strip() != "" :
            page.get_by_text("文案").click()
            page.locator("div.meta-input").click()
            page.locator("div.meta-input").fill(copywriting)#填文案
         
        #落地页   
        if pd.notna(landing) and str(landing).strip() != "" :
            page.locator("span").filter(has_text=re.compile(r"^落地页$")).click()
            if  re.match(r"^kfcapplinkurl",landing):
                page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^应用直达$")).click()
                page.locator("span.ellipsis.odc-text.odc-text-small.ellipsis.has-ellipsis").filter(has_text=re.compile(rf"^直达链接：{landing}$")).click(force=True)
                #page.locator("span.odc-tag-text.odc-text-mini").filter(has_text=re.compile(rf"^{dp_LINK}$")).click()
            else :
                page.locator("li.readonly.spaui-cursor-pointer").filter(has_text=re.compile(r"^微信小程序$")).click()
                page.locator("button.x-elem-add-btn.spaui-button.spaui-button-default.with-icon").first.click()
                page.get_by_role("textbox", name="gh 开头的小程序原始 ID").fill("gh_f01f85672b87")#肯德基小程序
                page.get_by_role("textbox", name="请输入小程序链接").fill(landing)
                page.get_by_role("button", name="新建至账户").click()
        
        #行动按钮
        if pd.notna(action_btn) and str(action_btn).strip() != "" :
            print("没写好~")
        
        
        #提交
        page.get_by_role("button", name="提交创意").click()
        page.get_by_role("button", name="返回创意管理").click()

        #wating
        page.get_by_role("button", name="编辑").first.wait_for(state="visible", timeout=300000)
        print(f"第{index+1}条创意 : 修改成功\n")

    input("所有创意修改成功，press Enter and quit")

def run_bili(playwright: Playwright,df) -> None:
    
    def extract_campaign_id(url):
        # 按 '/' 分割 URL，获取路径片段（忽略查询参数和锚点）
            path_fragments = url.split("/")
            # 遍历片段，找到 "campaign" 的位置
            for idx, fragment in enumerate(path_fragments):
                if fragment == "campaign":
                    # 取 "campaign" 下一个片段（即目标数字）
                    if idx + 1 < len(path_fragments):  # 避免索引越界，提升稳健性
                        return path_fragments[idx + 1]
            return None  # 未找到 campaign 片段时返回 None

    account_id = df.iloc[0, 35]# 广告账户
    print(f"创建账户:{account_id}\n")

    #begin
    try:
        chrome_exe_path = get_chromium_path()
        print(f"✅ 找到Chromium路径：{chrome_exe_path}\n")
        # 启动浏览器
        browser = playwright.chromium.launch(
            executable_path=chrome_exe_path,
            headless=False  # 显示浏览器窗口，方便排查
        )
        # 创建上下文（保留你的原有逻辑）
        context = browser.new_context(storage_state="auth_state_bili.json") if os.path.exists("auth_state_bili.json") else browser.new_context()
    except Exception as e:
        print(f"❌ 启动浏览器失败：{e}\n")
        sys.exit(1)

    #region登录
    #context = browser.new_context(storage_state="auth_state_bili.json") if os.path.exists("auth_state_bili.json") else browser.new_context()
    
    #小程序
    page1 = context.new_page()
    page = context.new_page()
    #"""
    page.goto(f"https://e.bilibili.com/site/account/select")
    """
    try:
        page.get_by_role("button", name="新建推广").wait_for(state="visible", timeout=300000)#阻塞，检测登录状态
        context.storage_state(path="auth_state_bili.json")#储存登录信息
    except Exception as e:
        page.get_by_role("textbox", name="请输入手机号").click()
        page.get_by_role("textbox", name="请输入手机号").fill("13761307177")
        #page.get_by_text("发送验证码").click()
        print("注意：必须在5分钟内登录，否则本次会话无效且强制关闭！")
        page.get_by_role("button", name="新建推广").wait_for(state="visible", timeout=300000)#阻塞，检测登录状态
        context.storage_state(path="auth_state_bili.json")#储存登录信息
    """
    input("确保当前已处于登录状态后，按下回车开始搭建！若未登录，Misa手机号：13761307177\n")
    context.storage_state(path="auth_state_bili.json")#储存登录信息
    page1.goto(f"https://ad.bilibili.com/#/assets/index?activeTab=my-small-game&type=list&account_id={account_id}")
    #"""
    #endregion
    
    #计划用
    plan_dict = {
    }

    PAGE_POSITION_MAP = {
            "信息流小卡_图片": "信息流小卡",
            "信息流大卡_图片": "信息流大卡",
            "信息流大卡_视频": "信息流大卡",
            "竖版视频流_视频": "竖屏视频流",
            "动态区信息流_视频": "动态区信息流",
            "横版视频": "信息流大卡"
        }

    #单元用

    for index,row in df.iterrows():

        #media表用
        #brand = row.iloc[3] #品牌
        strategy_ID = row.iloc[3] #策略ID
        campaign_NM = row.iloc[9] #活动名称
        #media = row.iloc[12] #媒体
        page_PST = row.iloc[14] #点位
        audience = row.iloc[18] #脱敏人群
        creative_NM = row.iloc[17] #创意名称
        city = row.iloc[20] #区域
        sell_Type = row.iloc[23] #购买类型
        mini_link = row.iloc[25] if pd.notna(row.iloc[25]) else "" #小程序链接
        mini_ID = row.iloc[27] if pd.notna(row.iloc[27]) else "" #小程序ID
        basic_LP = row.iloc[28] #打底落地页
        dp_LINK = row.iloc[29] if pd.notna(row.iloc[29]) else "" #deeplink
        imp_TLink = row.iloc[30] #曝光监测链接
        clk_Tlink = row.iloc[31] #点击监测链接
        cooperation1 = row.iloc[39] if 39 < len(row) else "" #合作协议1，如果没写就是空字符串
        cooperation2 = row.iloc[40] if 40 < len(row) else "" #合作协议2
    
        #创意用
        asset_NM = row.iloc[36] #图片或视频名称
        copywriting_TLE= row.iloc[37] #素材标题
        copywriting_DSC= row.iloc[38] #素材描述

        #region计划

        #添加小程序
        mini_NM = f"{strategy_ID}_{int(time.time())}"

        if not mini_link == "":

            page1.locator('span.dib.vm.ml8[data-v-3da72d1f]').filter(has_text=re.compile(r"^添加微信小程序$")).click()
            #page1.get_by_role("textbox", name="请填写小程序原始ID").click()
            page1.get_by_role("textbox", name="请填写小程序原始ID").fill(mini_ID)
            #page1.get_by_role("textbox", name="请填写小程序名称").click()
            page1.get_by_role("textbox", name="请填写小程序名称").fill(mini_NM)
            #page1.get_by_role("textbox", name="请填写小程序路径").click()
            page1.get_by_role("textbox", name="请填写小程序路径").fill(f"{mini_link}&trackid=__TRACKID__")

            page1.get_by_role("button", name="确定添加").click()
        
        plan_NM = f"{campaign_NM}_{audience}"

        if plan_NM in plan_dict :
            page.goto(f"https://ad.bilibili.com/#/promote/auto?campaign_id={plan_dict.get(plan_NM)}&account_id={account_id}")#计划内新建单元
        else :
            page.goto(f"https://ad.bilibili.com/#/promote/auto?type=1&account_id={account_id}")#新建计划
            
           #我知道了
            Iknow_div = page.locator('div.info-btn.fr[data-v-72c5a9f4]')
            try:
                Iknow_div.wait_for(state="attached", timeout=5000)
                if Iknow_div.count() > 0:
                    Iknow_div.click()
            except Exception:
                pass
            
            #内容种草
            page.locator('div.ppt-title[data-v-5eb66e25]').nth(1).click()
            
            #计划名称
            #page.get_by_role("textbox", name="请输入计划名称").click()
            page.get_by_role("textbox", name="请输入计划名称").fill(plan_NM)
            
            #计划预算
            #page.get_by_role("textbox", name="请输入不小于500，且只有2位小数").click()
            page.get_by_role("textbox", name="请输入不小于500，且只有2位小数").fill("500")
        #endregion

        #region单元
        unit_NM = f"{strategy_ID}_{campaign_NM}_{page_PST}_{creative_NM}_{audience}_{city}_{sell_Type}"

        #page.get_by_role("textbox", name="请输入单元名称").click()    
        page.get_by_role("textbox", name="请输入单元名称").fill(unit_NM)

        page.get_by_text("内容投放").click()

        if not mini_link == "":
            page.get_by_text("请选择微信小程序", exact=True).click()
            page.get_by_text(mini_NM, exact=True).click()
        else :
        #APP包
            page.get_by_text("请选择", exact=True).click()
            if sell_Type == "购买" : page.get_by_text("肯德基KFC-iOS").click() 
            else : page.get_by_text("肯德基KFC-安卓").click()

        #日期
        edit_time = page.get_by_role("link", name="编辑时段")

        try:
            edit_time.click(timeout=2000)
        except Exception as e:  # 核心修正：exception → Exception
            page.locator("span[data-v-53fbd318].vm").filter(has_text=re.compile(r"^展开关联产品、投放日期、频次和搜索快投等内容$")).click()#展开关联产品、投放日期、频次和搜索快投等内容
            edit_time.click()
        page.get_by_role("link", name="全部清除").click()
        page.get_by_role("button", name="确定").click()

        #出价
        page.get_by_text("CPM", exact=True).click()
        #page.get_by_role("textbox", name="请输入金额").click()
        page.get_by_role("textbox", name="请输入金额").fill("10")

        #单元日预算
        page.locator("#unit_budget_bid").get_by_text("指定日预算").click()
        #page.get_by_role("textbox", name="请输入不少于500，且只有2位小数").click()
        page.get_by_role("textbox", name="请输入不少于500，且只有2位小数").fill("500")

        #展示链接
        #page.get_by_role("textbox", name="请输入https链接开头的URL").first.click()
        page.get_by_role("textbox", name="请输入https链接开头的URL").first.fill(imp_TLink)
        

        #点击和播放3秒监控
        #page.get_by_role("textbox", name="请输入https链接开头的URL").nth(1).click()
        page.get_by_role("textbox", name="请输入https链接开头的URL").nth(1).fill(clk_Tlink)
        

        #点位
        time.sleep(0.25)#不知道要加载什么
        page.locator("span.ivu-switch.ivu-switch-small").nth(0).click()
        page.get_by_role("checkbox", name=PAGE_POSITION_MAP[page_PST]).check()
        if page_PST == "信息流大卡_视频":
            page.get_by_role("checkbox", name="动态区信息流").check() 
        elif page_PST == "动态区信息流_视频":
            page.get_by_role("checkbox", name="信息流大卡").check() #For疯四
        elif page_PST == "横版视频":
            page.get_by_role("checkbox", name="动态区信息流").check()
        #endregion
            
        #region新建创意

        #创意智能衍生 注意这块儿不要开启

        #添加图片/视频
        if page_PST in ["信息流小卡_图片","信息流大卡_图片"] :
            page.get_by_role("button", name="添加图片").click()
            page.get_by_role("textbox", name="请输入图片名称").click()
            page.get_by_role("textbox", name="请输入图片名称").fill(asset_NM)
            page.get_by_role("textbox", name="请输入图片名称").press("Enter")
            time.sleep(0.15)
            #选择图片
            #page.locator('div.asset-item-new[data-v-d464b5b8]').first.click()
            target_div = page.get_by_text(asset_NM, exact=True).first
            target_div.wait_for(state="visible", timeout=5000)
            div_box = target_div.bounding_box()
            click_x = div_box["x"] + div_box["width"] / 2  # 水平中心x坐标
            click_y = div_box["y"] - 50                   # 上边界向上50px的y坐标
            page.mouse.click(click_x, click_y)
            page.locator('button.ivu-btn.ivu-btn-primary.ok-btn[type="button"][data-v-26f6aad8] span').filter(has_text=re.compile(r"^确认$")).click()

        else :
            page.get_by_role("button", name="添加稿件/视频").click()
            page.get_by_role("link", name="我的视频").click()
            page.get_by_role("textbox", name="请输入视频名称搜索").click()
            page.get_by_role("textbox", name="请输入视频名称搜索").fill(asset_NM)
            page.get_by_role("textbox", name="请输入视频名称搜索").press("Enter")
            #page.get_by_text(asset_NM, exact=True).click()
            #page.locator('div.video-image[data-v-8640aa80]').first.click()
            page.locator('span.vm[data-v-8640aa80]').filter(has_text=asset_NM).first.click()
            page.locator('div.footer-actions button.ivu-btn.ivu-btn-primary.btn.primary[type="button"][data-v-788d48dc]').nth(0).click()
        
        #素材标题
        page.get_by_role("textbox", name="请输入2~40个字（移动场景建议18字以内）").click()
        page.get_by_role("textbox", name="请输入2~40个字（移动场景建议18字以内）").fill(copywriting_TLE)

        #唤起链接
        page.get_by_role("textbox", name="请输入唤起应用的链接").click()
        if not mini_link == "": 
            page.get_by_role("textbox", name="请输入唤起应用的链接").fill(mini_link)
        else:
            page.get_by_role("textbox", name="请输入唤起应用的链接").fill(dp_LINK)

        #素材描述
        page.get_by_role("textbox", name="请输入2 ~ 10个字，即客户端广告卡片中UP").click()
        page.get_by_role("textbox", name="请输入2 ~ 10个字，即客户端广告卡片中UP").fill(copywriting_DSC)
            
        #自定义落地页
        page.get_by_role("textbox", name="请使用https链接开头的URL", exact=True).click()
        page.get_by_role("textbox", name="请使用https链接开头的URL", exact=True).fill(basic_LP)

        #品牌头像
        if page.get_by_role("textbox", name="请选择品牌名称").is_visible() :
            page.get_by_role("textbox", name="请选择品牌名称").click()
            page.get_by_text("肯德基", exact=True).nth(0).click()
        
        #合作协议
        if pd.notna(cooperation1) and str(cooperation1).strip() != ""  :
            #page.get_by_text(r"^请选择$").click()
            page.locator('span.placeholder[data-v-ed7545cc]').filter(has_text=re.compile(r"^请选择$")).click()
            #page.locator('div.hp-poptip-trigger')
            page.locator('label.bd-checkbox').nth(cooperation1).click()
            
        if pd.notna(cooperation1) and str(cooperation1).strip() != "" and pd.notna(cooperation2) and str(cooperation2).strip() != ""  :
            #page.get_by_text("请选择补充资质").click()
            page.locator('label.bd-checkbox').nth(cooperation2).click()
            
        #endregion

        #保存
        time.sleep(0.5)
        page.get_by_role("button", name="保存").click()
        
        #存入计划ID
        page.get_by_role("button", name="新建创意").wait_for(state="visible", timeout=300000)
        if plan_NM not in plan_dict :
            plan_dict[plan_NM] = extract_campaign_id(page.url)
            #print(f"{plan_NM} : {plan_dict[plan_NM]}")
        print(f"第{index+1}条广告 : {unit_NM} 创建成功\n")

    input("广告创建完成，plz press enter and continue")
    context.close()
    browser.close()

def run_dy(playwright: Playwright,df) -> None:

    def get_url_param(url, param_name):
        """
        从URL中提取指定键名的查询参数值
        :param url: 完整的URL字符串(比如page.url返回的内容)
        :param param_name: 要提取的参数键名(比如"project_id"、"name"、"type")
        :return: 指定键名的参数值(字符串,无该参数则返回None;有多个值返回第一个)
        """
        # 步骤1：解析URL，拆分出查询参数部分（?后面的内容）
        parsed_url = urlparse(url)
        # 步骤2：解析查询参数为字典（key: [value1, value2,...]，自动处理重复参数）
        query_params = parse_qs(parsed_url.query)
        # 步骤3：提取指定键名的参数值（无则返回None，有则取第一个值）
        param_value = query_params.get(param_name, [None])[0]
        return param_value

    #begin
    try:
        chrome_exe_path = get_chromium_path()
        print(f"✅ 找到Chromium路径：{chrome_exe_path}\n")
        # 启动浏览器
        browser = playwright.chromium.launch(
            executable_path=chrome_exe_path,
            headless=False  # 显示浏览器窗口，方便排查
        )
        # 创建上下文（保留你的原有逻辑）
        context = browser.new_context(storage_state="auth_state_dy.json") if os.path.exists("auth_state_dy.json") else browser.new_context()
    except Exception as e:
        print(f"❌ 启动浏览器失败：{e}\n")
        sys.exit(1)

    #region登录
    page = context.new_page()
    #"""
    page.goto(f"https://business.oceanengine.com/site/account-manage/ad/bidding/superior/account")
    input("确保当前已处于登录状态后，按下回车开始搭建！\n")
    context.storage_state(path="auth_state_dy.json")#储存登录信息
    #"""
    #endregion

    for index,row in df.iterrows():

        #media表用
        #brand = row.iloc[3] #品牌
        strategy_ID = row.iloc[0] #策略ID
        campaign_NM = row.iloc[2] #活动名称
        #media = row.iloc[13] #媒体
        #page_PST = row.iloc[15] #点位
        audience = row.iloc[18] #脱敏人群
        creative_NM = row.iloc[20] #创意名称
        city = row.iloc[23] #区域
        sell_Type = row.iloc[26] #购买类型
        #basic_LP = row.iloc[27] #打底落地页
        #dp_LINK = row.iloc[29] #deeplink
        imp_TLink = row.iloc[35] #曝光监测链接
        clk_Tlink = row.iloc[36] #点击监测链接
        rta_ID = row.iloc[39]

        #copy用
        account_id = row.iloc[40]# 广告账户
        copy_AD = row.iloc[41] #复制的项目
        copy_UN = row.iloc[42] #复制的单元

        #前端自定义人群
        audience_MD = row.iloc[44] #媒体人群
        audience_tag = row.iloc[45] #人群类型

        print(f"创建账户:{account_id}")

        unit_NM = f"{strategy_ID}_{campaign_NM}_{creative_NM}_{rta_ID}_{audience_tag}_{audience}_{city}_{sell_Type}"
        #unit_NM = f"{strategy_ID}_{campaign_NM}_{media}_{page_PST}_{creative_NM}_{audience}_{city}_{sell_Type}_{rta_ID}_{audience_tag}"
        #unit_NM = f"{strategy_ID}_1月海绵宝宝_{creative_NM}_{audience_tag}_{audience}_{sell_Type}_{rta_ID}" #For彤姐
        #unit_NM = f"{strategy_ID}_1月 Kids海绵宝宝达推_{media}_{page_PST}_{creative_NM}_{audience}_{city}_{sell_Type}_{rta_ID}_{audience_tag}"#For雪宜
        
        page.goto(f"https://ad.oceanengine.com/superior/create-project?aadvid={account_id}&is_copy=1&project_id={copy_AD}")

        #region项目

        #首选媒体

        #自定义人群
        if pd.notna(audience_MD) and str(audience_MD).strip() != "" :
            page.get_by_text("自定义").nth(1).click()
            page.get_by_role("textbox", name="请输入", exact=True).click()
            page.get_by_role("textbox", name="请输入", exact=True).fill(audience_MD)
            page.get_by_role("button", name="定向").first.click() 

        #平台
        if sell_Type == "购买" or sell_Type == "追投2" : 
            page.locator('div[data-e2e="createproject_platformorientation_checkbox_group_component_1"]').click() #IOS
        else : 
            page.locator('div[data-e2e="createproject_platformorientation_checkbox_group_component_2"]').click() #Android
            page.locator('div[data-e2e="createproject_platformorientation_checkbox_group_component_3"]').click() #鸿蒙

        #指定时间和时间段根据给定模版


        #日预算 自己设

        #曝光监测链接
        page.get_by_role("textbox", name="请输入链接地址").first.click()
        page.get_by_role("textbox", name="请输入链接地址").first.fill(imp_TLink)

        #点击检测链接
        page.get_by_role("textbox", name="请输入链接地址").nth(1).click()
        page.get_by_role("textbox", name="请输入链接地址").nth(1).fill(clk_Tlink)

        #项目名称
        page.get_by_role("textbox", name="请输入项目名称").click()
        page.get_by_role("textbox", name="请输入项目名称").fill(unit_NM)

        #保存
        page.get_by_role("button", name="保存并关闭").click()
        
        #project_id
        page.get_by_role("button", name="项目工具").wait_for(state="visible", timeout=300000)
        project_id = get_url_param(page.url, "project_id")
        row.iloc[41] = project_id

        #endregion

        #region单元
        #copy_type=3 等于复制到别的项目
        page.goto(f"https://ad.oceanengine.com/superior/ads?aadvid={account_id}&is_copy=1&project_id={project_id}&campaign_type=1&ad_count=1&promotion_id={copy_UN}&copy_type=3") 
        #单元名称
        page.locator('textarea.ovui-textarea').wait_for(state="visible", timeout=300000)
        page.locator('textarea.ovui-textarea').click() 
        page.locator('textarea.ovui-textarea').fill(unit_NM)

        #保存
        page.get_by_role("button", name="保存并关闭").click()
        page.get_by_role("button", name="项目工具").wait_for(state="visible", timeout=300000)
        row.iloc[42] = ""
        
        #endregion
        
        #存入计划ID
        print(f"第{index+1}条广告 : {unit_NM} 创建成功\n")

    input("广告创建完成，plz press enter and continue\n")

    #输出新文件
    df = df.iloc[:, :44]
    df.to_excel("media_id_检查无误后可上传MOP.xlsx", index=False, engine="openpyxl")

    context.close()
    browser.close()

if __name__ == "__main__":

    df = read_excel_file()

    with sync_playwright() as playwright:

        total_cols = df.shape[1]  # df.shape[1] = 列数
        
        total_rows = df.shape[0]  # df.shape[0] = 行数
        
        if  total_cols >= 13 and df.columns[12] == "媒体":
            
            run_bili(playwright,df)

        elif total_cols >= 14 and (df.iloc[0, 13] == "抖音" or df.iloc[0, 13] == "番茄系媒体"):
            
            run_dy(playwright,df)

        elif total_cols >= 14 and (re.search(r'微信', df.iloc[0, 13]) or re.fullmatch(r'QQ|腾讯音乐|游戏', df.iloc[0, 13])):

            id_selector = None
            
            bitou_list = read_txt_file() if input("若搭建cpc广告，请输入大写的 Y 并 回车，否则直接 回车\n") == "Y" else None
            
            if bitou_list is not None:
            
                id_selector = IDCombinationSelector(bitou_list)#初始化避投包
            
            run_adq(playwright,df,id_selector)

        elif total_cols == 1 and total_rows == 0 :

            bitou_list = read_txt_file() 
            
            id_selector = IDCombinationSelector(bitou_list)#初始化避投包

            run_adq_cre_template(playwright,df,id_selector)
        
        else :

            run_adq_replace(playwright,df)
