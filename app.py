import streamlit as st
import openpyxl
from io import BytesIO
import os
import glob
import datetime
import pandas as pd
import tempfile
import time

# --- 打印 Word 附属文件执行函数 ---
def print_word_documents(main_id, spouse_id):
    doc_dir = os.path.join(os.getcwd(), "doc")
    if not os.path.exists(doc_dir):
        return False, "⚠️ 未在同级目录下找到 'doc' 文件夹，附属 Word 文档被跳过。"
        
    # 提取需要打印的有效身份证 (按人员角色独立入队)
    ids_to_process = []
    if main_id and str(main_id).strip():
        ids_to_process.append(("主借款人", str(main_id).strip()))
    if spouse_id and str(spouse_id).strip():
        ids_to_process.append(("配偶", str(spouse_id).strip()))
        
    if not ids_to_process:
        return False, "⚠️ 未填写任何有效身份证信息，授权书被跳过。"
        
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        
        # 启动隐藏且纯净的 Word 进程
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # 屏蔽弹窗
        word.Options.PrintBackground = False # 强制同步打印，防闪退
        
        # ✨ 核心修复：创建一个“替身”空白文档，强行给 Word 续命，防止它在第一份文档关闭后因为“0文档”而自杀！
        dummy_doc = word.Documents.Add()
        
        # 深度遍历替换助手：扫描文档正文、文本框、页眉等所有角落
        def replace_in_doc(doc_obj, find_str, replace_str):
            wdReplaceAll = 2
            wdFindContinue = 1
            for story in doc_obj.StoryRanges:
                try:
                    story.Find.Execute(FindText=find_str, MatchCase=False, MatchWholeWord=False, 
                                       MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, 
                                       Forward=True, Wrap=wdFindContinue, Format=False, 
                                       ReplaceWith=replace_str, Replace=wdReplaceAll)
                except Exception:
                    pass
                
                next_story = story.NextStoryRange
                while next_story:
                    try:
                        next_story.Find.Execute(FindText=find_str, MatchCase=False, MatchWholeWord=False, 
                                           MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, 
                                           Forward=True, Wrap=wdFindContinue, Format=False, 
                                           ReplaceWith=replace_str, Replace=wdReplaceAll)
                    except Exception:
                        pass
                    next_story = next_story.NextStoryRange

        # 锁定目标文件的绝对路径
        file1_path = os.path.abspath(os.path.join(doc_dir, "1_综合授权书.docx"))
        file2_path = os.path.abspath(os.path.join(doc_dir, "2_征信授权书.docx"))
        file3_path = os.path.abspath(os.path.join(doc_dir, "3_温馨提示.docx"))
        
        error_msgs = []

        # 按人头依次打印
        for role, id_num in ids_to_process:
            # --- 打印 1_综合授权书 ---
            if os.path.exists(file1_path):
                doc1 = None
                try:
                    # 使用具名参数 FileName 避免 Open 方法因为版本差异抛出 <unknown>.Open
                    doc1 = word.Documents.Open(FileName=file1_path, ReadOnly=False, Visible=False)
                    
                    # 极简替换逻辑：直接替换 idnumb
                    replace_in_doc(doc1, "idnumb", id_num)
                    
                    print(f"准备打印：1_综合授权书 ({role}: {id_num})")
                    doc1.PrintOut(Background=False)
                    time.sleep(1)
                except Exception as e:
                    err_str = f"1_综合授权书({role})失败: {e}"
                    print(err_str)
                    error_msgs.append(err_str)
                finally:
                    if doc1:
                        try: doc1.Close(SaveChanges=0) # 不保存关闭
                        except: pass

            # --- 打印 2_征信授权书 ---
            if os.path.exists(file2_path):
                doc2 = None
                try:
                    doc2 = word.Documents.Open(FileName=file2_path, ReadOnly=False, Visible=False)
                    
                    # 极简替换逻辑：直接替换 idnumb
                    replace_in_doc(doc2, "idnumb", id_num)
                    
                    print(f"准备打印：2_征信授权书 ({role}: {id_num})")
                    doc2.PrintOut(Background=False)
                    time.sleep(1)
                except Exception as e:
                    err_str = f"2_征信授权书({role})失败: {e}"
                    print(err_str)
                    error_msgs.append(err_str)
                finally:
                    if doc2:
                        try: doc2.Close(SaveChanges=0)
                        except: pass
                        
        # --- 打印 3_温馨提示 (不分人头，总共只打一份) ---
        if os.path.exists(file3_path):
            doc3 = None
            try:
                # 提示文件无需替换，以纯只读模式安全打开
                doc3 = word.Documents.Open(FileName=file3_path, ReadOnly=True, Visible=False)
                print("准备打印：3_温馨提示")
                doc3.PrintOut(Background=False)
                time.sleep(1)
            except Exception as e:
                err_str = f"3_温馨提示失败: {e}"
                print(err_str)
                error_msgs.append(err_str)
            finally:
                if doc3:
                    try: doc3.Close(SaveChanges=0)
                    except: pass

        try:
            # 打扫战场：关掉替身文档，并彻底退出 Word
            if dummy_doc:
                dummy_doc.Close(SaveChanges=0)
            word.Quit()
            del word
        except:
            pass
        pythoncom.CoUninitialize()
        
        if error_msgs:
            uniq_errs = list(dict.fromkeys(error_msgs))
            return False, "部分附属文档可能失败：\n" + "\n".join(uniq_errs)
            
        return True, "✅ 附属 Word 授权书及提示文件已全部完美打印完毕！"
        
    except ImportError:
        return False, "⚠️ 缺少 pywin32 库，无法自动控制 Word 打印。"
    except Exception as e:
        try:
            if 'word' in locals() and word:
                word.Quit()
            pythoncom.CoUninitialize()
        except:
            pass
        return False, f"⚠️ Word 打印宏观控制出错: {e}"

# --- 打印 Excel 核心执行函数 ---
def print_excel_worksheets(file_bytes, filename, marital_status, provident_fund_loan):
    unique_str = str(int(time.time() * 1000))
    safe_filename = f"print_{unique_str}_{filename}"
    temp_path = os.path.join(tempfile.gettempdir(), safe_filename)
    
    with open(temp_path, "wb") as f:
        f.write(file_bytes)
    
    try:
        import win32com.client
        import pythoncom  
        
        pythoncom.CoInitialize() 
        
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False 
        
        wb = excel.Workbooks.Open(temp_path)
        
        has_provident_fund = False
        if provident_fund_loan:
            val_str = str(provident_fund_loan).strip()
            if val_str not in ["", "0", "0.0", "无"]:
                try:
                    pf_val = float(val_str.replace('万', '').replace('元', '').replace(',', ''))
                    if pf_val > 0:
                        has_provident_fund = True
                except ValueError:
                    has_provident_fund = True

        for ws in wb.Sheets:
            if ws.Visible == -1: 
                sheet_name = ws.Name
                copies = 1 
                # print(sheet_name)
                
                if sheet_name == "信息录入":
                    continue
                elif sheet_name == "单身声明":
                    if marital_status == "已婚":
                        continue 
                    else:
                        copies = 3 
                elif sheet_name == "抵押合同":
                    copies = 2
                elif sheet_name == "个人贷款申请表-公积金":
                    if not has_provident_fund:
                        continue 
                    else:
                        copies = 1
                
                try:
                    print(f"正在打印：{sheet_name} (份数: {copies})")
                    ws.PrintOut(Copies=copies)
                    time.sleep(1) 
                except Exception as sheet_err:
                    print(f"工作表 {sheet_name} 打印出现小状况: {sheet_err}")
                    
        wb.Close(False)
        excel.Quit()
        
        del wb
        del excel
        
        pythoncom.CoUninitialize() 
        
        return True, "✅ Excel 智能打印任务已成功发送！"
        
    except ImportError:
        try:
            os.startfile(temp_path, "print")
            return True, "✅ 已调用系统默认打印机打印 Excel。"
        except Exception as ex:
            return False, f"调用系统打印失败: {ex}"
    except Exception as e:
        try:
            if 'wb' in locals(): 
                wb.Close(False)
                del wb
            if 'excel' in locals(): 
                excel.Quit()
                del excel
            import pythoncom
            pythoncom.CoUninitialize()
        except:
            pass
        return False, f"Excel 打印出错，请检查打印机状态: {e}"

# --- 页面基本配置 ---
st.set_page_config(page_title="房贷信息录入及打印系统 V1.5", layout="wide", page_icon="📝")

st.title("📝 房贷信息录入及打印系统 V1.5")

# --- 第一步：选择本地模板 ---
st.header("1. 选择本地Excel模板")

excel_files = glob.glob("*.xlsx")

if not excel_files:
    st.warning("⚠️ 当前目录下未检测到任何 .xlsx 模板文件！请将模板文件放入与 app.py 相同的文件夹中并刷新页面。")
    template_file = None
else:
    template_file = st.selectbox("检测到以下模板，请选择：", excel_files)

st.divider()

# --- 读取本地的 account.csv ---
account_options = ["二手房"]
account_df = None
if os.path.exists("account.csv"):
    try:
        account_df = pd.read_csv("account.csv", encoding='gbk').dropna(how='all')
        valid_projects = account_df['楼盘名称'].dropna().astype(str).unique().tolist()
        account_options.extend(valid_projects)
    except Exception as e:
        st.warning(f"读取 account.csv 失败: {e}")

def on_account_change():
    selected = st.session_state.account_selector
    if selected != "二手房" and account_df is not None:
        row = account_df[account_df['楼盘名称'].astype(str) == selected].iloc[0]
        st.session_state.project_address = str(row['地址']) if pd.notna(row['地址']) else ""
        st.session_state.project_name = str(row['收款人']) if pd.notna(row['收款人']) else ""
        
        acc_val = row['收款帐号']
        if pd.notna(acc_val):
            if isinstance(acc_val, float) and acc_val.is_integer():
                st.session_state.collection_account = str(int(acc_val))
            else:
                st.session_state.collection_account = str(acc_val)
        else:
            st.session_state.collection_account = ""

# --- 第二步：填写表单 ---
st.header("2. 填写字段信息")
st.subheader("👤 人员基础信息")
col1, col2 = st.columns(2)
with col1:
    st.markdown("**👉 主借款人 (第2行)**")
    main_name = st.text_input("姓名 (B2)")
    main_id = st.text_input("身份证 (D2)")
    main_gender = st.selectbox("性别 (C2)", ["", "男", "女"])
    main_phone = st.text_input("电话 (F2)")
    main_education = st.selectbox("学历 (J2)", ["", "初中", "高中", "大专", "本科", "硕士"])
    main_residence = st.text_input("户籍 (K2)")
    st.markdown("<br>", unsafe_allow_html=True)
    home_phone = st.text_input("家庭固话 (B21)", help="若无固话可留空")
    
with col2:
    st.markdown("**👉 主借款人配偶 (第3行)**")
    spouse_name = st.text_input("姓名 (B3)")
    spouse_id = st.text_input("身份证 (D3)")
    spouse_gender = st.selectbox("性别 (C3)", ["", "男", "女"])
    spouse_phone = st.text_input("电话 (F3)")
    spouse_education = st.selectbox("学历 (J3)", ["", "初中及以下", "高中/中专", "大专", "本科", "硕士及以上"])
    spouse_residence = st.text_input("户籍 (K3)")
    
st.markdown("---")

st.subheader("🏠 房产与资金信息")
st.selectbox("📌 快速选择楼盘 (自动填充下方地址、收款人及账号)", 
             options=account_options, 
             key="account_selector", 
             on_change=on_account_change)
project_address = st.text_input("地址 (B23)", key="project_address")
col3, col4, col5 = st.columns(3)

with col3:
    total_price = st.text_input("购房总价 (B14)")
    down_payment = st.text_input("首期 (B15)")   
with col4:
    building_area = st.text_input("建筑面积 (B12)")
    internal_area = st.text_input("套内面积 (B13)")
with col5:
    project_name = st.text_input("收款人 (B24)", key="project_name")
    collection_account = st.text_input("收款帐号 (B25)", key="collection_account")

st.markdown("---")

st.subheader("💼 贷款与工作信息")
col6, col7, col8 = st.columns(3) 

with col6:
    st.markdown("**👉 贷款基本信息**")
    provident_fund_loan = st.text_input("公积金贷款金额 (F10)")
    commercial_loan = st.text_input("商业贷款金额 (G10)")
    repayment_method = st.text_input("还款方式 (F12)")
    repayment_card = st.text_input("还款卡号 (F13)")
    marital_status = st.selectbox("婚姻状况 (F14)", ["", "未婚", "已婚", "离异", "丧偶"])
    
with col7:
    st.markdown("**👉 借款人工作信息**")
    borrower_company = st.text_input("单位名称 (F15)")
    borrower_address = st.text_input("工作地址 (F16)")
    borrower_title = st.text_input("职务 (F17)")
    borrower_work_phone = st.text_input("单位固话 (F18)")
    borrower_work_years = st.text_input("工作年限 (F19)")
    income_1 = st.text_input("月收入1 (B19)")

with col8:
    st.markdown("**👉 配偶工作信息**")
    spouse_company = st.text_input("单位名称 (F20)")
    spouse_address = st.text_input("工作地址 (F21)")
    spouse_title = st.text_input("职务 (F22)")
    spouse_work_phone = st.text_input("单位固话 (F23)")
    spouse_work_years = st.text_input("工作年限 (F24)")
    income_2 = st.text_input("月收入2 (B20)")

st.markdown("<br>", unsafe_allow_html=True)

with st.form("main_form"):
    submitted = st.form_submit_button("✅ 填写完毕，生成Excel文件", use_container_width=True)

# --- 第三步：数据处理与保存 ---
if submitted:
    if not template_file:
        st.error("⚠️ 请先确保当前目录下存在有效的Excel模板文件！")
    else:
        try:
            fill_data = {
                "B2": main_name, "C2": main_gender, "D2": main_id, "F2": main_phone, 
                "J2": main_education, "K2": main_residence,
                "B3": spouse_name, "C3": spouse_gender, "D3": spouse_id, "F3": spouse_phone, 
                "J3": spouse_education, "K3": spouse_residence,
                "B21": home_phone,
                "B24": project_name, "B23": project_address, "B25": collection_account,
                "B12": building_area, "B13": internal_area, "B14": total_price, "B15": down_payment,
                "F10": provident_fund_loan, "G10": commercial_loan, 
                "F12": repayment_method, "F13": repayment_card, "F14": marital_status,
                "F15": borrower_company, "F16": borrower_address, "F17": borrower_title, 
                "F18": borrower_work_phone, "F19": borrower_work_years, "B19": income_1,
                "F20": spouse_company, "F21": spouse_address, "F22": spouse_title, 
                "F23": spouse_work_phone, "F24": spouse_work_years, "B20": income_2
            }

            wb = openpyxl.load_workbook(template_file)
            sheet = wb.active
            error_cells = [] 

            for coord, val in fill_data.items():
                if val and str(val).strip() != "":
                    try:
                        sheet[coord] = val
                    except AttributeError as e:
                        if "MergedCell" in str(e) or "read-only" in str(e):
                            error_cells.append(coord)
                        else:
                            raise e
            
            if error_cells:
                st.error(f"⚠️ **生成中断**！检测到以下坐标为合并单元格的非首个坐标，无法写入：{', '.join(error_cells)}")
            else:
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                today_date = datetime.date.today().strftime("%Y%m%d") 
                borrower_name = main_name.strip() if main_name.strip() else "未命名"
                dynamic_filename = f"{today_date}-{borrower_name}.xlsx"
                
                # 存入 Session State
                st.session_state['generated_excel'] = output.getvalue()
                st.session_state['generated_filename'] = dynamic_filename
                
                # 提取打印相关参数入 Session
                st.session_state['print_marital_status'] = marital_status
                st.session_state['print_provident_fund_loan'] = provident_fund_loan
                st.session_state['print_main_id'] = main_id
                st.session_state['print_spouse_id'] = spouse_id

        except Exception as e:
            st.error(f"处理Excel文件时发生不可预知的错误: {e}")

# --- 第四步：展示下载与打印按钮 ---
if st.session_state.get('generated_excel') is not None:
    st.success("🎉 Excel 文件已成功生成！请选择下载保存或直接打印。")
    st.info("💡 **【关于双面打印的特别提示】** 👉 **请您务必在电脑的 [设置] -> [打印机和扫描仪] -> [打印首选项] 中，将该打印机的默认模式手动设为“双面打印”**。", icon="🖨️")

    
    col_dl, col_pr = st.columns(2)
    
    with col_dl:
        st.download_button(
            label="📥 下载填写好的 Excel 文件",
            data=st.session_state['generated_excel'],
            file_name=st.session_state['generated_filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    with col_pr:
        if st.button("🖨️ 直接打印", type="primary", use_container_width=True):
            with st.spinner('正在执行 Excel 及 Word 文档的全流程打印任务，请耐心稍候...'):
                # 1. 首先打印 Excel
                success_ex, msg_ex = print_excel_worksheets(
                    st.session_state['generated_excel'], 
                    st.session_state['generated_filename'],
                    st.session_state.get('print_marital_status', ''),
                    st.session_state.get('print_provident_fund_loan', '')
                )

                if success_ex:
                    st.toast(msg_ex)
                    
                    # 2. Excel 打印成功后，继续打印 Word 附属文件
                    success_wd, msg_wd = print_word_documents(
                        st.session_state.get('print_main_id', ''),
                        st.session_state.get('print_spouse_id', '')
                    )
                    
                    if success_wd:
                        st.success(f"{msg_ex} 并且 {msg_wd}")
                    else:
                        st.warning(f"Excel 打印成功，但附属文档有提示：\n{msg_wd}")
                else:
                    st.error(msg_ex)