import subprocess
import time
from playwright.sync_api import sync_playwright
import atexit
import pyautogui

# 全局变量，用于管理浏览器和自动化相关的资源
chrome_process = None  # Chrome 浏览器进程
playwright = None     # Playwright 实例
browser = None        # CDP 浏览器连接
context = None        # 浏览器上下文，包含标签页等信息
page = None          # 当前操作的标签页

def launch_browser():
    """启动 Chrome 浏览器，并开启远程调试功能"""
    global chrome_process
    chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
    debugging_port = "--remote-debugging-port=9955"  # 指定调试端口
    
    # 启动浏览器，使用 subprocess 以便能够控制进程
    command = f'"{chrome_path}" {debugging_port}'
    chrome_process = subprocess.Popen(command, shell=True)
    time.sleep(2)  # 等待浏览器完全启动

def close_browser():
    """优雅地关闭浏览器及相关资源，确保资源正确释放，避免出现浏览器未正确关闭的提示"""
    global chrome_process, playwright, browser, context, page
    try:
        if browser and browser.is_connected():
            try:
                # 获取所有标签页并逐个关闭
                pages = context.pages
                print(f"Found {len(pages)} pages to close")
                
                # 从后往前关闭标签页，避免影响索引
                for p in reversed(pages):
                    try:
                        p.close()
                        print(f"Closed page {pages.index(p) + 1}")
                        time.sleep(2)  # 每个标签页关闭后短暂等待
                    except Exception as e:
                        print(f"Error closing page {pages.index(p) + 1}: {e}")
                
                print("All pages closed, sleeping 5 seconds...")
                time.sleep(5)  # 等待所有标签页完全关闭
            except Exception as e:
                print(f"Error during pages cleanup: {e}")

            # 关闭浏览器上下文和 CDP 连接
            if context:
                try:
                    context.close()
                    print("Context closed")
                except Exception as e:
                    print(f"Error closing context: {e}")
            
            try:
                browser.close()
                print("Browser connection closed")
            except Exception as e:
                print(f"Error closing browser: {e}")

        # 停止 Playwright
        if playwright:
            try:
                playwright.stop()
                print("Playwright stopped")
            except Exception as e:
                print(f"Error stopping playwright: {e}")

        # 终止 Chrome 进程
        if chrome_process:
            try:
                chrome_process.terminate()
                print("Chrome process terminated")
                time.sleep(2)  # 等待进程完全终止
            except Exception as e:
                print(f"Error terminating chrome process: {e}")

    except Exception as e:
        print(f"Error during browser cleanup: {e}")
    finally:
        # 清理全局变量，避免资源残留
        chrome_process = None
        playwright = None
        browser = None
        context = None
        page = None

def get_perplexity_answer(question):
    """获取 Perplexity AI 的回答，包括连接浏览器、提问和获取答案的完整流程"""
    global playwright, browser, context, page
    
    try:
        # 首次建立 CDP 连接或重新连接
        if not browser or not browser.is_connected():
            print("Establishing initial CDP connection...")
            browser = playwright.chromium.connect_over_cdp("http://localhost:9955")
            context = browser.contexts[0]  # 获取现有上下文
            pages = context.pages  # 获取所有标签页
            if pages:
                page = pages[0]  # 使用现有标签页
                print("Connected to existing page")
            else:
                print("No pages found, creating new one")
                page = context.new_page()
            
        # 访问 Perplexity AI 网站
        print(f"Navigating to Perplexity AI...")
        page.goto("https://www.perplexity.ai")
        time.sleep(2)  # 等待页面加载
        
        # 输入问题并发送
        print(f"Typing question: {question}")
        page.keyboard.type(question)
        time.sleep(1)
        page.keyboard.press("Enter")
        
        # 等待 AI 生成回答
        print("Waiting for answer to load (20 seconds)...")
        time.sleep(20)
        
        # 断开 CDP 连接以便发送浏览器层面快捷键，因为cdp时浏览器拒绝其他程序访问，尝试了很多方法都不行，只能断开了
        print("Disconnecting CDP...")
        if browser and browser.is_connected():
            if context:
                context.close()  # 关闭所有上下文
            browser.close()
        browser = None
        context = None
        page = None
        time.sleep(2)  # 等待连接完全断开
        
        # 使用 pyautogui 发送快捷键，安装的拓展名Save my Chatbot……，快捷键是我在chrome拓展页面自己设置的，也可以修改成其他的
        print("Sending Alt+Shift+I hotkey...")
        try:
            pyautogui.hotkey('alt', 'shift', 'i')
            print("Hotkey sent successfully")
        except Exception as e:
            print(f"Error sending hotkey: {e}")
            raise
        time.sleep(2)  # 等待插件响应
        
        # 重新建立 CDP 连接以便后续操作
        print("Reconnecting CDP...")
        try:
            browser = playwright.chromium.connect_over_cdp("http://localhost:9955")
            context = browser.contexts[0]  # 获取现有上下文
            pages = context.pages  # 获取所有标签页
            if pages:
                page = pages[0]  # 使用现有标签页
                print("Reconnected to existing page")
            else:
                print("No pages found, creating new one")
                page = context.new_page()
            time.sleep(2)  # 等待连接建立
        except Exception as e:
            print(f"Error reconnecting CDP: {e}")
            raise
        
        return "Answer has been downloaded by the browser extension"
            
    except Exception as e:
        print(f"Error in get_perplexity_answer: {str(e)}")
        return f"发生错误: {str(e)}"

def main(question):
    """主函数：协调整个自动化流程的执行"""
    global playwright
    
    try:
        # 启动浏览器
        launch_browser()
        time.sleep(2)
        
        # 启动 Playwright
        playwright = sync_playwright().start()
        
        # 获取回答
        result = get_perplexity_answer(question)
        return result
        
    finally:
        # 确保资源被正确关闭
        close_browser()

# 注册程序退出时的回调函数，确保在程序异常退出时也能清理资源
atexit.register(close_browser)

if __name__ == "__main__":
    # 测试问题
    changequestionhere = "What are the most famous news in 2024?"
    result = main(changequestionhere)
    print("\nQuestion:", changequestionhere)
    print("\nStatus:", result)

