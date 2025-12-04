import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import NoSuchElementException

# ==================== 配置区域 ====================
# 您可以在这里修改所有配置参数

# Edge WebDriver路径（根据您的实际情况修改）
EDGE_DRIVER_PATH = "F:/Python_Project/QRCode/edgedriver/msedgedriver.exe"

# 目标网页URL（根据您的需求修改）
TARGET_URL = "https://v.wjx.cn/vm/eNt5SO6.aspx"  # 测试用网页，请替换为您需要的网页

# 刷新参数
MAX_REFRESH_RETRIES = 20  # 最大刷新尝试次数
REFRESH_INTERVAL = 3  # 刷新间隔时间（秒）

# 字典：输入框上方附近的中文文字 -> 要填写的内容
# 请根据您遇到的实际网页修改这个字典
INPUT_MAPPING_DICT = {
    "姓名": "杜松煜",
    "名字": "杜松煜",
    "学院": "软件学院",
    "班级": "计算机类2512",
    "学号": "20254536",
    "电话": "15869519789",
    "联系方式": "15869519789"
}


# ==================== 核心类 ====================

class EdgeAutoFiller:
    def __init__(self, driver_path):
        """初始化Edge浏览器驱动"""
        print("正在初始化Edge浏览器...")

        # Edge浏览器选项
        edge_options = Options()
        edge_options.add_argument('--disable-blink-features=AutomationControlled')
        edge_options.add_argument('--disable-gpu')
        edge_options.add_argument('--start-maximized')
        edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        edge_options.add_experimental_option('useAutomationExtension', False)

        # 创建Edge服务
        service = Service(executable_path=driver_path)

        # 初始化浏览器
        self.driver = webdriver.Edge(service=service, options=edge_options)
        self.driver.set_page_load_timeout(30)
        print("✓ Edge浏览器初始化成功")

    def open_and_find_inputs(self, url):
        """打开网页并持续刷新直到找到输入框"""
        print(f"打开网页: {url}")

        for attempt in range(1, MAX_REFRESH_RETRIES + 1):
            print(f"尝试 #{attempt}/{MAX_REFRESH_RETRIES}")

            try:
                self.driver.get(url)
                time.sleep(2)  # 等待页面加载

                input_elements = self.find_input_elements()

                if input_elements:
                    print(f"✓ 找到 {len(input_elements)} 个输入框")
                    return input_elements
                else:
                    print("未找到输入框，准备刷新...")

                    if attempt < MAX_REFRESH_RETRIES:
                        time.sleep(REFRESH_INTERVAL)
                        self.driver.refresh()

            except Exception as e:
                print(f"尝试 #{attempt} 失败: {e}")
                if attempt < MAX_REFRESH_RETRIES:
                    time.sleep(REFRESH_INTERVAL)

        print(f"✗ 经过 {MAX_REFRESH_RETRIES} 次尝试仍未找到输入框")
        return []

    def find_input_elements(self):
        """查找页面中所有可见的输入框"""
        input_elements = []

        # 查找各种输入框
        input_selectors = [
            'input[type="text"]',
            'input[type="password"]',
            'input[type="email"]',
            'input[type="number"]',
            'input[type="tel"]',
            'input[type="search"]',
            'input[type="url"]',
            'textarea',
            'input:not([type])',
        ]

        for selector in input_selectors:
            try:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for element in elements:
                    if element.is_displayed() and element.is_enabled():
                        input_elements.append(element)
            except:
                continue

        return input_elements

    def extract_chinese_near_input(self, input_element):
        """提取输入框上方附近的中文文字"""
        try:
            # 方法1: 查找前面的label元素
            for i in range(1, 6):
                try:
                    xpath = f".//preceding::label[{i}]"
                    label = input_element.find_element(By.XPATH, xpath)
                    text = label.text.strip()
                    if text and self.contains_chinese(text):
                        return text
                except:
                    pass

            # 方法2: 查找前面的兄弟元素
            for i in range(1, 6):
                try:
                    xpath = f".//preceding-sibling::*[{i}]"
                    sibling = input_element.find_element(By.XPATH, xpath)
                    text = sibling.text.strip()
                    if text and self.contains_chinese(text):
                        return text
                except:
                    pass

            # 方法3: 查找父元素中的文本
            for i in range(1, 4):
                try:
                    xpath = f".//ancestor::*[{i}]"
                    ancestor = input_element.find_element(By.XPATH, xpath)
                    text = ancestor.text.strip()
                    lines = text.split('\n')
                    for line in lines:
                        line = line.strip()
                        if line and self.contains_chinese(line) and len(line) < 50:
                            return line
                except:
                    continue

            # 方法4: 检查placeholder
            placeholder = input_element.get_attribute("placeholder")
            if placeholder and self.contains_chinese(placeholder):
                return placeholder

            return ""

        except Exception as e:
            print(f"提取文本时出错: {e}")
            return ""

    def contains_chinese(self, text):
        """检查是否包含中文"""
        if not text:
            return False
        chinese_pattern = re.compile(r'[\u4e00-\u9fff]')
        return bool(chinese_pattern.search(text))

    def fill_inputs_using_dict(self, input_elements):
        """使用字典自动填写输入框"""
        print(f"开始填写输入框，字典大小: {len(INPUT_MAPPING_DICT)}")

        filled_count = 0
        for i, input_element in enumerate(input_elements):
            try:
                chinese_text = self.extract_chinese_near_input(input_element)

                if chinese_text:
                    print(f"输入框 #{i + 1}: 找到文本 '{chinese_text}'")

                    for key, value in INPUT_MAPPING_DICT.items():
                        if key in chinese_text:
                            input_element.clear()
                            input_element.send_keys(value)
                            print(f"  ✓ 填写: '{value}' (匹配: '{key}')")
                            filled_count += 1
                            break
                    else:
                        print(f"  ⚠ 未找到匹配项")
                else:
                    print(f"输入框 #{i + 1}: 未找到附近中文文本")

            except Exception as e:
                print(f"输入框 #{i + 1} 填写失败: {e}")

        print(f"填写完成，共填写 {filled_count} 个输入框")
        return filled_count

    def find_and_click_button(self):
        """查找并点击第一个可点击的按钮"""
        print("查找页面中的按钮...")

        try:
            # 查找所有可能的按钮
            buttons = []

            # <button>元素
            buttons.extend(self.driver.find_elements(By.TAG_NAME, "button"))

            # <input type="button">等
            button_types = ['submit', 'button', 'reset']
            for btn_type in button_types:
                selector = f'input[type="{btn_type}"]'
                buttons.extend(self.driver.find_elements(By.CSS_SELECTOR, selector))

            # 具有按钮角色的元素
            buttons.extend(self.driver.find_elements(By.CSS_SELECTOR, '[role="button"]'))

            # 查找可点击的按钮
            clickable_buttons = []
            for btn in buttons:
                try:
                    if btn.is_displayed() and btn.is_enabled():
                        clickable_buttons.append(btn)
                except:
                    continue

            print(f"找到 {len(clickable_buttons)} 个可点击按钮")

            # 点击第一个按钮
            if clickable_buttons:
                first_button = clickable_buttons[0]
                button_text = first_button.text.strip()
                if not button_text:
                    button_text = first_button.get_attribute("value") or "无文本按钮"

                print(f"尝试点击按钮: '{button_text}'")

                # 记录点击前状态
                before_url = self.driver.current_url
                before_title = self.driver.title

                # 点击按钮
                first_button.click()
                time.sleep(3)  # 等待页面响应

                # 记录点击后状态
                after_url = self.driver.current_url
                after_title = self.driver.title

                print(f"✓ 按钮点击成功")
                print(f"  点击前 - 标题: {before_title}, URL: {before_url}")
                print(f"  点击后 - 标题: {after_title}, URL: {after_url}")
                print(f"  页面是否变化: {'是' if before_url != after_url else '否'}")

                return True
            else:
                print("未找到可点击的按钮")
                return False

        except Exception as e:
            print(f"点击按钮时出错: {e}")
            return False

    def get_page_info(self):
        """获取当前页面信息"""
        info = {}

        try:
            info["标题"] = self.driver.title
            info["URL"] = self.driver.current_url

            # 查找页面中的中文文本
            page_source = self.driver.page_source
            chinese_pattern = re.compile(r'[\u4e00-\u9fff]{2,}')
            chinese_matches = chinese_pattern.findall(page_source)

            info["中文短语数量"] = len(chinese_matches)

            # 获取前5个中文短语
            unique_chinese = list(set(chinese_matches))
            info["中文示例"] = unique_chinese[:5]

        except Exception as e:
            info["错误"] = str(e)

        return info


# ==================== 主程序 ====================

def main():
    """主函数：执行自动化任务"""
    print("=" * 60)
    print("Edge浏览器自动化程序 - 开始执行")
    print("=" * 60)

    # 显示配置信息
    print(f"目标网页: {TARGET_URL}")
    print(f"映射字典: {INPUT_MAPPING_DICT}")
    print(f"最大刷新次数: {MAX_REFRESH_RETRIES}")
    print(f"刷新间隔: {REFRESH_INTERVAL}秒\n")

    # 创建自动化对象
    automator = EdgeAutoFiller(EDGE_DRIVER_PATH)

    try:
        # 1. 打开网页并查找输入框
        print("阶段1: 查找输入框")
        input_elements = automator.open_and_find_inputs(TARGET_URL)

        if not input_elements:
            print("未找到输入框，程序结束")
            return

        # 2. 根据字典自动填写
        print("\n阶段2: 自动填写输入框")
        filled_count = automator.fill_inputs_using_dict(input_elements)

        # 3. 查找并点击按钮
        print("\n阶段3: 查找并点击按钮")
        button_clicked = automator.find_and_click_button()

        # 4. 获取页面信息
        print("\n阶段4: 获取页面信息")
        page_info = automator.get_page_info()

        print("\n页面信息:")
        for key, value in page_info.items():
            if key == "中文示例":
                print(f"  {key}:")
                for i, text in enumerate(value, 1):
                    print(f"    {i}. {text}")
            else:
                print(f"  {key}: {value}")

        # 5. 显示任务总结
        print("\n" + "=" * 60)
        print("任务完成总结:")
        print(f"  输入框找到数量: {len(input_elements)}")
        print(f"  输入框成功填写: {filled_count}")
        print(f"  按钮点击成功: {'是' if button_clicked else '否'}")
        print(f"  页面标题: {page_info.get('标题', 'N/A')}")
        print("=" * 60)

        # 保持浏览器打开，用户可手动查看结果
        print("\n注意: 浏览器将保持打开状态，请手动关闭")
        print("您可以查看页面并验证填写结果")

        # 等待用户手动关闭
        while True:
            try:
                time.sleep(1)
            except KeyboardInterrupt:
                print("\n检测到键盘中断，程序结束")
                break

    except KeyboardInterrupt:
        print("\n\n用户中断程序")
    except Exception as e:
        print(f"\n程序执行出错: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()