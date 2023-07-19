import win32com.client

def copy_selected_text():
    # 创建Word应用程序对象
    word_app = win32com.client.Dispatch("Word.Application")

    # 显示Word应用程序窗口（可根据需要进行调整）
    word_app.Visible = True

    try:
        # 获取当前活动文档对象
        doc = word_app.ActiveDocument

        # 获取当前选中的文本（如果存在）
        if doc.Application.Selection.Type == win32com.client.constants.wdSelectionIP:
            selected_text = doc.Application.Selection.Range.Text
            print("选中的文本：", selected_text)

            # 将选中的文本复制到剪贴板
            doc.Application.Selection.Copy()
        else:
            print("没有选中的文本")
    except Exception as e:
        print("复制选中的文本时出现错误：", str(e))
    finally:
        # 退出Word应用程序
        word_app.Quit()

# 调用函数复制选中的文本
copy_selected_text()