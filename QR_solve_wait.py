"""
Edge浏览器网页自动化程序 - 精简确认版
功能：打开网页，检测输入框，根据字典自动填写，识别并点击提交按钮
"""

import time
import random
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.action_chains import ActionChains

# ==================== 配置区域 ====================

# Edge WebDriver路径，请根据电脑本地的edgedriver地址修改
EDGE_DRIVER_PATH = ""

# 目标网页URL，请根据问卷的URL进行修改，若问卷为二维码可使用https://qrcodereader.net/翻译为URL
TARGET_URL = "" #

# 刷新参数
MAX_REFRESH_RETRIES = 15
REFRESH_INTERVAL = 0.5

# 字典：输入框上方附近的中文文字 -> 要填写的内容
INPUT_MAPPING_DICT = {
    "学校": "test1",
    "姓名": "test2",
    "名字": "test3",
    "学院": "test4",
    "班级": "test5",
    "学号": "test6",
    "电话": "test7",
    "联系方式": "test8",
    "寝室": "test9"
}


# ==================== 核心类 ====================

class EdgeAutoFiller:
    def __init__(self, driver_path):
        """初始化Edge浏览器驱动 - 增强反检测"""
        print("正在初始化Edge浏览器...")

        edge_options = Options()

        # 基础设置
        edge_options.add_argument('--start-maximized')
        edge_options.add_argument('--disable-gpu')
        edge_options.add_argument('--no-sandbox')

        # 反自动化检测设置
        edge_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        edge_options.add_experimental_option('useAutomationExtension', False)

        # 禁用自动化控制特征
        edge_options.add_argument('--disable-blink-features=AutomationControlled')

        # 添加用户代理和语言设置
        edge_options.add_argument(
            '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0')
        edge_options.add_argument('--lang=zh-CN')

        # 禁用一些可能被检测的特征
        edge_options.add_argument('--disable-web-security')
        edge_options.add_argument('--allow-running-insecure-content')
        edge_options.add_argument('--disable-dev-shm-usage')

        # 禁用自动控制提示
        prefs = {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "profile.default_content_setting_values.notifications": 2,
            "excludeSwitches": ["enable-automation"],
            "useAutomationExtension": False
        }
        edge_options.add_experimental_option("prefs", prefs)

        service = Service(executable_path=driver_path)
        self.driver = webdriver.Edge(service=service, options=edge_options)
        self.driver.set_page_load_timeout(30)

        # 执行JavaScript代码来隐藏自动化特征
        self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': '''
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                });
                Object.defineProperty(navigator, 'plugins', {
                    get: () => [1, 2, 3, 4, 5]
                });
                Object.defineProperty(navigator, 'languages', {
                    get: () => ['zh-CN', 'zh']
                });
            '''
        })

        print("✓ Edge浏览器初始化成功")

    def random_delay(self, min_time=0.5, max_time=1.0):
        """随机延迟，模拟人类操作"""
        delay = random.uniform(min_time, max_time)
        time.sleep(delay)

    def human_like_typing(self, element, text):
        """模拟人类打字，逐个字符输入"""
        element.clear()
        for char in text:
            element.send_keys(char)
            # self.random_delay(0.05, 0.2)  # 随机延迟模拟打字间隔

    def open_webpage(self, url):
        """打开网页"""
        print(f"打开网页: {url}")
        try:
            self.driver.get(url)
            time.sleep(0.2)  # 固定等待页面加载，加快速度
            print("✓ 网页打开成功")
            return True
        except Exception as e:
            print(f"打开网页失败: {e}")
            return False

    def refresh_webpage(self):
        """刷新当前网页"""
        try:
            print("刷新网页...")
            self.driver.refresh()
            time.sleep(REFRESH_INTERVAL)  # 固定等待，加快速度
            print("✓ 网页刷新成功")
            return True
        except Exception as e:
            print(f"刷新网页失败: {e}")
            return False

    def find_inputs_with_retry(self):
        """持续刷新直到找到输入框"""
        print("开始查找输入框...")

        for attempt in range(1, MAX_REFRESH_RETRIES + 1):
            print(f"尝试 #{attempt}/{MAX_REFRESH_RETRIES}")

            try:
                input_elements = self.find_input_elements()

                if input_elements:
                    print(f"✓ 找到 {len(input_elements)} 个输入框")
                    return input_elements
                else:
                    print("未找到输入框，准备刷新...")

                    if attempt < MAX_REFRESH_RETRIES:
                        self.refresh_webpage()

            except Exception as e:
                print(f"尝试 #{attempt} 失败: {e}")
                if attempt < MAX_REFRESH_RETRIES:
                    time.sleep(REFRESH_INTERVAL)

        print(f"✗ 经过 {MAX_REFRESH_RETRIES} 次尝试仍未找到输入框")
        return []

    def find_input_elements(self):
        """查找页面中所有可见的输入框"""
        input_elements = []

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
                    try:
                        if element.is_displayed() and element.is_enabled():
                            input_elements.append(element)
                    except:
                        continue
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

            # 方法5: 查找前面的文本节点
            try:
                xpath = f".//preceding::text()[normalize-space()][last()]"
                script = f"""
                var input = arguments[0];
                var xpath = '{xpath}';
                var iterator = document.evaluate(xpath, input, null, XPathResult.ANY_TYPE, null);
                var node = iterator.iterateNext();
                return node ? node.textContent.trim() : '';
                """
                text = self.driver.execute_script(script, input_element)
                if text and self.contains_chinese(text):
                    return text
            except:
                pass

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
        unfilled_inputs = []

        for i, input_element in enumerate(input_elements):
            try:
                chinese_text = self.extract_chinese_near_input(input_element)

                if chinese_text:
                    print(f"输入框 #{i + 1}: 找到文本 '{chinese_text}'")

                    matched = False
                    for key, value in INPUT_MAPPING_DICT.items():
                        if key in chinese_text:
                            # 模拟人类点击和打字
                            ActionChains(self.driver).move_to_element(input_element).click().perform()
                            time.sleep(0.1)  # 极短延迟

                            # 模拟人类打字
                            self.human_like_typing(input_element, value)

                            print(f"  ✓ 填写: '{value}' (匹配: '{key}')")
                            filled_count += 1
                            matched = True
                            break

                    if not matched:
                        print(f"  ⚠ 未找到匹配项")
                        unfilled_inputs.append({
                            'index': i + 1,
                            'text': chinese_text,
                            'element': input_element
                        })
                else:
                    print(f"输入框 #{i + 1}: 未找到附近中文文本")
                    unfilled_inputs.append({
                        'index': i + 1,
                        'text': '未找到文本',
                        'element': input_element
                    })

            except Exception as e:
                print(f"输入框 #{i + 1} 填写失败: {e}")
                unfilled_inputs.append({
                    'index': i + 1,
                    'text': f'错误: {str(e)[:30]}',
                    'element': input_element
                })

        # 尝试填写未匹配的输入框
        if unfilled_inputs:
            print(f"\n尝试填写未匹配的输入框...")
            for info in unfilled_inputs:
                try:
                    ActionChains(self.driver).move_to_element(info['element']).click().perform()
                    time.sleep(0.1)  # 极短延迟
                    self.human_like_typing(info['element'], "默认填写")
                    filled_count += 1
                    print(f"输入框 #{info['index']}: 已填写默认值")
                except:
                    print(f"输入框 #{info['index']}: 无法填写默认值")

        print(f"填写完成，共填写 {filled_count} 个输入框")
        return filled_count, len(input_elements)

    def check_inputs_filled(self, input_elements):
        """检查所有输入框是否已填写"""
        all_filled = True
        unfilled_list = []

        for i, input_element in enumerate(input_elements):
            try:
                value = input_element.get_attribute("value")
                if not value or value.strip() == "":
                    all_filled = False
                    unfilled_list.append(i + 1)
            except:
                all_filled = False
                unfilled_list.append(i + 1)

        if all_filled:
            print("✓ 所有输入框均已填写")
        else:
            print(f"⚠ 以下输入框未填写: {unfilled_list}")

        return all_filled

    def find_submit_button(self):
        """查找提交按钮"""
        print("查找提交按钮...")

        # 主要尝试通过ID查找
        try:
            button = self.driver.find_element(By.ID, "ctlNext")
            if button.is_displayed() and button.is_enabled():
                return self.get_button_info(button, "ID: ctlNext")
        except:
            pass

        # 通过类名查找
        try:
            button = self.driver.find_element(By.CLASS_NAME, "submitbtn")
            if button.is_displayed() and button.is_enabled():
                return self.get_button_info(button, "CLASS: submitbtn")
        except:
            pass

        # 通过XPath查找
        try:
            button = self.driver.find_element(By.XPATH, "//div[contains(text(), '提交')]")
            if button.is_displayed() and button.is_enabled():
                return self.get_button_info(button, "XPATH: //div[contains(text(), '提交')]")
        except:
            pass

        print("未找到提交按钮")
        return None

    def get_button_info(self, button_element, selector_used):
        """获取按钮的详细信息"""
        button_info = {
            'element': button_element,
            'tag_name': button_element.tag_name,
            'text': button_element.text.strip(),
            'selector_used': selector_used,
            'attributes': {},
        }

        try:
            attributes = ['id', 'class', 'type', 'value', 'name']
            for attr in attributes:
                value = button_element.get_attribute(attr)
                if value:
                    button_info['attributes'][attr] = value
        except:
            pass

        return button_info

    def click_submit_button(self, button_info):
        """点击提交按钮 - 简化版，不读取页面信息"""
        button_element = button_info['element']

        print(f"准备点击按钮: '{button_info['text']}'")
        print(
            f"按钮信息: 标签={button_info['tag_name']}, ID={button_info['attributes'].get('id', '无')}, 类名={button_info['attributes'].get('class', '无')}")
        print(f"使用选择器: {button_info['selector_used']}")

        try:
            # 记录点击前状态（仅记录，不详细显示）
            before_url = self.driver.current_url

            # 滚动到按钮位置
            self.driver.execute_script("arguments[0].scrollIntoView();", button_element)

            # 点击前等待1秒（不打印）
            time.sleep(0.5)

            print("正在点击按钮...")

            # 使用ActionChains点击
            ActionChains(self.driver).move_to_element(button_element).click().perform()

            print("✓ 按钮点击成功")

            # 等待页面响应（简化等待）
            time.sleep(2)

            # 记录点击后状态（仅记录，不显示）
            after_url = self.driver.current_url

            return {
                'success': True,
                'page_changed': before_url != after_url,
            }

        except Exception as e:
            print(f"点击按钮时出错: {e}")
            # 尝试JavaScript点击
            try:
                print("尝试使用JavaScript点击...")
                self.driver.execute_script("arguments[0].click();", button_element)
                print("JavaScript点击成功")
                return {
                    'success': True,
                    'page_changed': False
                }
            except Exception as e2:
                print(f"JavaScript点击也失败: {e2}")
                return {
                    'success': False,
                    'error': str(e)
                }

    def find_and_click_submit_button(self, input_elements):
        """查找并点击提交按钮 - 自动执行，无用户确认"""
        print("开始查找并点击提交按钮...")

        # 检查输入框是否已填写（仅检查，不影响提交）
        all_filled = self.check_inputs_filled(input_elements)

        if not all_filled:
            print("⚠ 部分输入框未填写，但仍尝试提交...")

        # 查找提交按钮
        button_info = self.find_submit_button()

        if button_info:
            print("✓ 找到提交按钮")

            # 点击提交按钮
            click_result = self.click_submit_button(button_info)

            return {
                'button_found': True,
                'button_clicked': click_result['success'],
                'button_info': button_info,
                'click_result': click_result
            }
        else:
            print("未找到提交按钮")
            return {
                'button_found': False,
                'button_clicked': False,
                'error': '未找到提交按钮'
            }


# ==================== 主程序 ====================

def main():
    """主函数：执行自动化任务"""
    print("=" * 50)
    print("Edge浏览器自动化程序 - 精简确认版")
    print("=" * 50)

    print(f"目标网页: {TARGET_URL}")
    print(f"映射字典: {INPUT_MAPPING_DICT}")
    print(f"最大刷新次数: {MAX_REFRESH_RETRIES}")
    print(f"刷新间隔: {REFRESH_INTERVAL}秒\n")

    # 创建自动化对象
    automator = EdgeAutoFiller(EDGE_DRIVER_PATH)

    try:
        # 1. 打开网页
        print("阶段1: 打开网页")
        if not automator.open_webpage(TARGET_URL):
            print("打开网页失败，程序结束")
            return

        # 只在打开网页后进行用户确认
        print("\n网页已打开，是否继续执行自动化任务？")
        user_confirm = input("输入 'y' 或 'yes' 继续，其他键退出: ").strip().lower()

        if user_confirm not in ['y', 'yes']:
            print("用户选择退出程序")
            return

        # 2. 查找输入框（带重试）
        print("阶段2: 查找输入框")
        input_elements = automator.find_inputs_with_retry()

        if not input_elements:
            print("未找到输入框，程序结束")
            return

        # 3. 根据字典自动填写
        print("\n阶段3: 自动填写输入框")
        filled_count, total_inputs = automator.fill_inputs_using_dict(input_elements)

        # 4. 查找并点击提交按钮（自动执行，无用户确认）
        print("\n阶段4: 查找并点击提交按钮")
        button_result = automator.find_and_click_submit_button(input_elements)

        # 显示按钮识别结果（简化显示）
        print("\n按钮识别结果:")
        if button_result['button_found']:
            btn_info = button_result['button_info']
            print(f"  找到按钮: 是")
            print(f"  按钮文本: '{btn_info['text']}'")
            print(f"  按钮点击: {'成功' if button_result['button_clicked'] else '失败'}")
        else:
            print(f"  找到按钮: 否")
            print(f"  错误信息: {button_result.get('error', '未知错误')}")

        # 5. 显示任务总结
        print("\n" + "=" * 50)
        print("任务完成总结:")
        print(f"  输入框总数: {total_inputs}")
        print(f"  成功填写数: {filled_count}")
        print(f"  提交按钮点击: {'成功' if button_result.get('button_clicked') else '失败'}")
        print("=" * 50)

        # 保持浏览器打开
        print("\n浏览器保持打开状态，请手动关闭...")
        print("10秒后自动退出程序")
        time.sleep(10)

    except KeyboardInterrupt:
        print("\n用户中断程序")
    except Exception as e:
        print(f"\n程序执行出错: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()