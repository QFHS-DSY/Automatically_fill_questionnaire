"""
Edge浏览器网页自动化程序 - 增强按钮识别版
功能：打开网页，检测输入框，根据字典自动填写，识别并点击自定义按钮
"""

import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.common.exceptions import NoSuchElementException

# ==================== 配置区域 ====================
# 您可以在这里修改所有配置参数

# Edge WebDriver路径，请根据电脑本地的edgedriver地址修改
EDGE_DRIVER_PATH = ""

# 目标网页URL，请根据问卷的URL进行修改，若问卷为二维码可使用https://qrcodereader.net/翻译为URL
TARGET_URL = "" #

# 刷新参数
MAX_REFRESH_RETRIES = 15  # 最大刷新尝试次数
REFRESH_INTERVAL = 2  # 刷新间隔时间（秒）

# 字典：输入框上方附近的中文文字 -> 要填写的内容
# 请根据问卷内容修改这个字典
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
                # 使用XPath查找前面的文本节点
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
        unfilled_inputs = []  # 记录未填写的输入框信息

        for i, input_element in enumerate(input_elements):
            try:
                chinese_text = self.extract_chinese_near_input(input_element)

                if chinese_text:
                    print(f"输入框 #{i + 1}: 找到文本 '{chinese_text}'")

                    matched = False
                    for key, value in INPUT_MAPPING_DICT.items():
                        if key in chinese_text:
                            input_element.clear()
                            input_element.send_keys(value)
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
                    # 尝试直接填写默认值
                    info['element'].clear()
                    info['element'].send_keys("默认填写")
                    filled_count += 1
                    print(f"输入框 #{info['index']}: 已填写默认值")
                except:
                    print(f"输入框 #{info['index']}: 无法填写默认值")

        print(f"填写完成，共填写 {filled_count} 个输入框")
        return filled_count, len(input_elements)

    def check_inputs_filled(self, input_elements):
        """检查所有输入框是否已填写"""
        print("检查输入框是否已填写...")

        all_filled = True
        unfilled_list = []

        for i, input_element in enumerate(input_elements):
            try:
                value = input_element.get_attribute("value")
                if not value or value.strip() == "":
                    all_filled = False
                    unfilled_list.append(i + 1)
                    print(f"输入框 #{i + 1}: 未填写")
                else:
                    print(f"输入框 #{i + 1}: 已填写 '{value}'")
            except Exception as e:
                all_filled = False
                unfilled_list.append(i + 1)
                print(f"输入框 #{i + 1}: 检查失败 - {e}")

        if all_filled:
            print("✓ 所有输入框均已填写")
        else:
            print(f"⚠ 以下输入框未填写: {unfilled_list}")

        return all_filled

    def find_submit_buttons(self):
        """查找页面中所有可能的提交按钮，返回详细信息"""
        print("查找页面中的提交按钮...")

        # 查找按钮的各种可能性
        button_selectors = [
            # 根据您提供的示例: <div id="ctlNext" class="submitbtn mainBgColor">提交</div>
            'div.submitbtn',
            'div[class*="submitbtn"]',
            'div[class*="submit"]',
            'div[class*="btn"]',
            'div[id*="submit"]',
            'div[id*="next"]',
            'div[id*="ctlNext"]',
            'div[id*="ctl_Next"]',
            'button.submitbtn',
            'button[class*="submitbtn"]',
            'input[type="submit"]',
            'input[value*="提交"]',
            'input[value*="下一步"]',
            'button[type="submit"]',
            'a[class*="submitbtn"]',
            'a[class*="btn"]',
            'span[class*="submitbtn"]',
            'span[class*="btn"]',
        ]

        all_buttons = []
        found_buttons = []

        for selector in button_selectors:
            try:
                buttons = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if buttons:
                    print(f"选择器 '{selector}' 找到 {len(buttons)} 个元素")

                    for button in buttons:
                        try:
                            if button.is_displayed() and button.is_enabled():
                                # 获取按钮的详细信息
                                button_info = self.get_button_info(button, selector)
                                all_buttons.append(button_info)

                                # 检查是否是提交按钮（通过文本或属性）
                                button_text = button_info['text']
                                if button_text and (
                                        "提交" in button_text or "下一步" in button_text or "确认" in button_text or "Submit" in button_text.lower()):
                                    found_buttons.append(button_info)
                                    print(f"  发现提交按钮: '{button_text}'")
                                elif not button_text:
                                    # 检查value属性
                                    value = button_info['attributes'].get('value', '')
                                    if "提交" in value or "下一步" in value or "确认" in value or "submit" in value.lower():
                                        found_buttons.append(button_info)
                                        print(f"  发现提交按钮 (通过属性): '{value}'")
                        except Exception as e:
                            print(f"  处理按钮时出错: {e}")
                            continue
            except Exception as e:
                print(f"使用选择器 '{selector}' 时出错: {e}")
                continue

        # 如果没有找到，尝试通过文本查找
        if not found_buttons:
            print("通过CSS选择器未找到提交按钮，尝试通过文本查找...")
            try:
                # 查找包含"提交"、"下一步"等文本的元素
                text_xpaths = [
                    "//*[contains(text(), '提交')]",
                    "//*[contains(text(), '下一步')]",
                    "//*[contains(text(), '确认')]",
                    "//*[contains(text(), 'Submit')]",
                    "//*[contains(text(), 'submit')]",
                ]

                for xpath in text_xpaths:
                    try:
                        elements = self.driver.find_elements(By.XPATH, xpath)
                        for element in elements:
                            try:
                                if element.is_displayed() and element.is_enabled():
                                    button_info = self.get_button_info(element, f"XPath: {xpath}")
                                    found_buttons.append(button_info)
                                    print(f"通过文本找到按钮: '{button_info['text']}'")
                            except:
                                continue
                    except:
                        continue
            except Exception as e:
                print(f"通过文本查找按钮时出错: {e}")

        # 返回找到的所有按钮信息
        return {
            'all_buttons': all_buttons,
            'submit_buttons': found_buttons
        }

    def get_button_info(self, button_element, selector_used):
        """获取按钮的详细信息"""
        button_info = {
            'element': button_element,
            'tag_name': button_element.tag_name,
            'text': button_element.text.strip(),
            'selector_used': selector_used,
            'attributes': {},
            'location': {},
            'size': {},
        }

        try:
            # 获取常用属性
            attributes = ['id', 'class', 'type', 'value', 'name', 'role', 'aria-label', 'onclick']
            for attr in attributes:
                value = button_element.get_attribute(attr)
                if value:
                    button_info['attributes'][attr] = value

            # 获取位置和大小
            location = button_element.location
            size = button_element.size
            button_info['location'] = {'x': location.get('x', 0), 'y': location.get('y', 0)}
            button_info['size'] = {'width': size.get('width', 0), 'height': size.get('height', 0)}

        except Exception as e:
            print(f"获取按钮信息时出错: {e}")

        return button_info

    def click_submit_button(self, button_info, delay_before_click=1):
        """点击提交按钮，点击前有延时"""
        button_element = button_info['element']

        print(f"准备点击按钮: '{button_info['text']}'")
        print(
            f"按钮信息: 标签={button_info['tag_name']}, ID={button_info['attributes'].get('id', '无')}, 类名={button_info['attributes'].get('class', '无')}")

        try:
            # 记录点击前状态
            before_url = self.driver.current_url
            before_title = self.driver.title

            print(f"点击前等待 {delay_before_click} 秒...")
            time.sleep(delay_before_click)  # 点击前延时

            print("正在点击按钮...")

            # 使用JavaScript点击，避免一些点击拦截
            self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                                       button_element)
            time.sleep(0.5)  # 等待滚动完成

            # 方法1: 使用JavaScript点击
            self.driver.execute_script("arguments[0].click();", button_element)

            print("✓ 按钮点击成功")

            # 等待页面响应
            time.sleep(3)

            # 记录点击后状态
            after_url = self.driver.current_url
            after_title = self.driver.title

            print(f"点击前 - 标题: {before_title}, URL: {before_url}")
            print(f"点击后 - 标题: {after_title}, URL: {after_url}")
            print(f"页面是否变化: {'是' if before_url != after_url else '否'}")

            return {
                'success': True,
                'page_changed': before_url != after_url,
                'before_click': {'url': before_url, 'title': before_title},
                'after_click': {'url': after_url, 'title': after_title}
            }

        except Exception as e:
            print(f"点击按钮时出错: {e}")

            # 尝试备用点击方法
            try:
                print("尝试备用点击方法...")
                self.driver.execute_script(
                    "arguments[0].dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, view: window}));",
                    button_element)
                print("备用点击方法成功")

                time.sleep(3)

                return {
                    'success': True,
                    'page_changed': False,
                    'before_click': {'url': self.driver.current_url, 'title': self.driver.title},
                    'after_click': {'url': self.driver.current_url, 'title': self.driver.title}
                }
            except Exception as e2:
                print(f"备用点击方法也失败: {e2}")

                return {
                    'success': False,
                    'error': str(e),
                    'page_changed': False
                }

    def find_and_click_submit_button(self, input_elements):
        """查找并点击提交按钮，在所有输入框填写完成后"""
        print("查找提交按钮...")

        # 首先检查所有输入框是否已填写
        all_filled = self.check_inputs_filled(input_elements)

        if not all_filled:
            print("部分输入框未填写，尝试补充填写...")
            # 尝试重新填写未填写的输入框
            for i, input_element in enumerate(input_elements):
                try:
                    value = input_element.get_attribute("value")
                    if not value or value.strip() == "":
                        # 尝试填写默认值
                        input_element.clear()
                        input_element.send_keys("默认填写")
                        print(f"补充填写输入框 #{i + 1}")
                except:
                    pass

            # 再次检查
            all_filled = self.check_inputs_filled(input_elements)

        # 如果输入框已填写，查找并点击按钮
        if all_filled:
            print("输入框已全部填写，开始查找提交按钮...")

            # 查找所有可能的提交按钮
            buttons_result = self.find_submit_buttons()

            if buttons_result['submit_buttons']:
                print(f"找到 {len(buttons_result['submit_buttons'])} 个可能的提交按钮")

                # 显示所有找到的提交按钮
                for i, button_info in enumerate(buttons_result['submit_buttons']):
                    print(f"\n按钮 #{i + 1}:")
                    print(f"  文本: '{button_info['text']}'")
                    print(f"  标签: {button_info['tag_name']}")
                    print(f"  ID: {button_info['attributes'].get('id', '无')}")
                    print(f"  类名: {button_info['attributes'].get('class', '无')}")
                    print(f"  选择器: {button_info['selector_used']}")

                # 尝试点击第一个提交按钮
                first_button = buttons_result['submit_buttons'][0]
                click_result = self.click_submit_button(first_button, delay_before_click=1)

                return {
                    'buttons_found': len(buttons_result['submit_buttons']),
                    'button_clicked': True,
                    'click_result': click_result,
                    'button_info': first_button
                }
            else:
                print("未找到提交按钮，尝试点击第一个可点击的按钮...")

                # 如果没有找到明显的提交按钮，尝试点击第一个可点击的按钮
                if buttons_result['all_buttons']:
                    print(f"找到 {len(buttons_result['all_buttons'])} 个按钮，尝试点击第一个")

                    first_button = buttons_result['all_buttons'][0]
                    click_result = self.click_submit_button(first_button, delay_before_click=1)

                    return {
                        'buttons_found': len(buttons_result['all_buttons']),
                        'button_clicked': True,
                        'click_result': click_result,
                        'button_info': first_button
                    }
                else:
                    print("未找到任何可点击的按钮")
                    return {
                        'buttons_found': 0,
                        'button_clicked': False,
                        'click_result': {'success': False, 'error': '未找到按钮'}
                    }
        else:
            print("输入框未全部填写，不点击提交按钮")
            return {
                'buttons_found': 0,
                'button_clicked': False,
                'click_result': {'success': False, 'error': '输入框未全部填写'}
            }

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

            # 查找表单数量
            forms = self.driver.find_elements(By.TAG_NAME, "form")
            info["表单数量"] = len(forms)

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
        filled_count, total_inputs = automator.fill_inputs_using_dict(input_elements)

        # 3. 查找并点击提交按钮
        print("\n阶段3: 查找并点击提交按钮")
        button_result = automator.find_and_click_submit_button(input_elements)

        # 显示按钮识别结果
        print("\n按钮识别结果:")
        print(f"  找到按钮总数: {button_result['buttons_found']}")
        print(f"  按钮是否点击: {button_result['button_clicked']}")

        if button_result.get('button_info'):
            btn_info = button_result['button_info']
            print(f"  点击的按钮信息:")
            print(f"    文本: '{btn_info['text']}'")
            print(f"    标签: {btn_info['tag_name']}")
            print(f"    ID: {btn_info['attributes'].get('id', '无')}")
            print(f"    类名: {btn_info['attributes'].get('class', '无')}")

        if button_result.get('click_result'):
            click_res = button_result['click_result']
            print(f"  点击结果: {'成功' if click_res.get('success') else '失败'}")
            if click_res.get('page_changed'):
                print(f"  页面跳转: 是")
                print(f"  跳转前URL: {click_res.get('before_click', {}).get('url', '未知')}")
                print(f"  跳转后URL: {click_res.get('after_click', {}).get('url', '未知')}")
            else:
                print(f"  页面跳转: 否")

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
        print(f"  输入框总数: {total_inputs}")
        print(f"  成功填写数: {filled_count}")
        print(f"  提交按钮点击: {'成功' if button_result.get('button_clicked') else '失败'}")

        if button_result.get('click_result') and button_result['click_result'].get('page_changed'):
            print(f"  页面跳转: 是")
        else:
            print(f"  页面跳转: 否")

        print(f"  最终页面标题: {page_info.get('标题', 'N/A')}")
        print("=" * 60)

        # 保持浏览器打开，用户可手动查看结果
        print("\n注意: 浏览器将保持打开状态，请手动关闭")
        print("您可以查看页面并验证填写结果")

        # 等待用户手动关闭
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n检测到键盘中断，程序结束")

    except KeyboardInterrupt:
        print("\n\n用户中断程序")
    except Exception as e:
        print(f"\n程序执行出错: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()