#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票-销售数据合并工具 - Streamlit GUI
"""

import streamlit as st
import tempfile
import os
import io
import zipfile
from merge import process_merge

st.set_page_config(page_title="发票-销售数据合并工具", page_icon="📊", layout="centered")

# 自定义样式
st.markdown("""
<style>
    .stMainBlockContainer { max-width: 720px; }
    div[data-testid="stFileUploader"] { margin-bottom: 0; }
    .upload-ok { color: #28a745; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.markdown("## 📊 发票-销售数据合并工具")
st.markdown("上传发票和销售 Excel 文件，自动完成精确 + 模糊匹配合并。")

st.divider()

# ── 文件上传 ──
st.markdown("#### 1. 选择文件")

with st.expander("📖 文件格式要求（点击查看）"):
    st.markdown("""
**发票文件（销项）** 要求：
- 必须包含以下列：`货物或应税劳务名称`、`数量`、`金额`、`税率`
- 商品名称支持带分类前缀格式（如 `*蔬菜*西红柿`），程序会自动去除前缀

**销售文件** 要求包含以下 Sheet 及对应列：

| Sheet 名称 | 必需列 | 开票列（程序填写） | 匹配税率范围 |
|:---:|:---:|:---:|:---:|
| 蔬菜 | 商品名称 | 开票数量、开票金额 | 13%、免税 |
| 肉蛋 | 商品名称 | 开票数量、开票金额 | 9%、13%、免税 |
| 9% | 商品名称 | 已开票数量、已开票金额 | 9% |
| 13% | 商品名称 | 已开票数量、已开票金额 | 13% |

> 不存在的 Sheet 会自动跳过，不影响其他 Sheet 处理。
""")

invoice_file = st.file_uploader("发票文件（销项 .xlsx）", type=["xlsx", "xls"], key="inv")
sales_file = st.file_uploader("销售文件（.xlsx）", type=["xlsx", "xls"], key="sal")

# 上传状态提示
if invoice_file and sales_file:
    st.info(f"已选择：**{invoice_file.name}** + **{sales_file.name}**", icon="✅")

# ── 高级设置 ──
with st.expander("⚙️ 高级设置"):
    threshold = st.slider(
        "模糊匹配阈值",
        min_value=0.50, max_value=1.00, value=0.75, step=0.05,
        help="相似度 >= 此值才算匹配。越高越严格，越低匹配越多但可能不准确。"
    )

st.divider()

# ── 合并 ──
st.markdown("#### 2. 开始处理")

if st.button("🚀 开始合并", type="primary",
             disabled=not (invoice_file and sales_file),
             use_container_width=True):

    with tempfile.TemporaryDirectory() as tmpdir:
        invoice_path = os.path.join(tmpdir, invoice_file.name)
        sales_path = os.path.join(tmpdir, sales_file.name)

        with open(invoice_path, 'wb') as f:
            f.write(invoice_file.getvalue())
        with open(sales_path, 'wb') as f:
            f.write(sales_file.getvalue())

        output_dir = os.path.join(tmpdir, "output")

        progress_bar = st.progress(0, text="准备中...")

        def on_progress(step, total, msg):
            progress_bar.progress(step / total, text=msg)

        try:
            result = process_merge(
                invoice_path=invoice_path,
                sales_path=sales_path,
                output_dir=output_dir,
                threshold=threshold,
                progress_callback=on_progress,
            )

            progress_bar.progress(1.0, text="完成！")
            stats = result['stats']

            # ── 结果 ──
            st.divider()
            st.markdown("#### 3. 处理结果")
            st.success("处理完成！")

            c1, c2, c3 = st.columns(3)
            c1.metric("精确匹配", f"{stats['exact_matches']} 条")
            c2.metric("模糊匹配", f"{stats['fuzzy_matches']} 条")
            c3.metric("未匹配发票", f"{stats['unmatched_invoices']} 条")

            # 各 Sheet 详情
            with st.expander("📋 各 Sheet 匹配详情", expanded=False):
                for sheet_name, s in stats['sheet_stats'].items():
                    total = s['total']
                    matched = s['exact_match'] + s['fuzzy_match']
                    pct = f"{matched / total * 100:.1f}%" if total else "0%"
                    st.markdown(
                        f"- **{sheet_name}**：精确 {s['exact_match']} / "
                        f"模糊 {s['fuzzy_match']} / "
                        f"未匹配 {s['unmatched']}（匹配率 {pct}）"
                    )

            # ── 下载 ──
            st.divider()
            st.markdown("#### 4. 下载结果")

            # 收集所有输出文件
            output_files = {}
            sales_basename = os.path.splitext(os.path.basename(sales_path))[0]

            with open(result['output_file'], 'rb') as f:
                fname = f"{sales_basename}_已更新.xlsx"
                output_files[fname] = f.read()

            if result['fuzzy_file']:
                with open(result['fuzzy_file'], 'rb') as f:
                    output_files["模糊匹配详情.xlsx"] = f.read()

            if result['unmatched_file']:
                with open(result['unmatched_file'], 'rb') as f:
                    output_files["未匹配记录.xlsx"] = f.read()

            # 打包 ZIP
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for name, data in output_files.items():
                    zf.writestr(name, data)
            zip_buf.seek(0)

            # 一键下载全部
            st.download_button(
                "📦 一键下载全部（ZIP）",
                data=zip_buf.getvalue(),
                file_name="合并结果.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

            # 单独下载
            with st.expander("或单独下载各文件"):
                xlsx_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                for fname, data in output_files.items():
                    st.download_button(
                        f"📥 {fname}",
                        data=data,
                        file_name=fname,
                        mime=xlsx_mime,
                        use_container_width=True,
                    )

        except Exception as e:
            progress_bar.empty()
            st.error(f"处理失败：{e}")
            import traceback
            with st.expander("错误详情"):
                st.code(traceback.format_exc())
