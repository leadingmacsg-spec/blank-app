import streamlit as st

import pandas as pd

import io



st.set_page_config(page_title="Excel费率处理系统", layout="wide")



st.title("📊 Excel费率处理系统")

st.write("上传总表和分表Excel文件，自动计算费率并进行匹配合并")



# 创建两列用于上传文件

col1, col2 = st.columns(2)



with col1:

    st.subheader("总表")

    total_file = st.file_uploader("上传总表Excel文件", type=['xlsx', 'xls'], key='total')



with col2:

    st.subheader("分表")

    sub_file = st.file_uploader("上传分表Excel文件", type=['xlsx', 'xls'], key='sub')



# 处理上传的文件

if total_file and sub_file:

    try:

        # 读取Excel文件

        total_df = pd.read_excel(total_file)

        sub_df = pd.read_excel(sub_file)

       

        st.success("✅ 文件上传成功！")

       

        # 显示原始数据

        with st.expander("查看原始数据"):

            st.subheader("总表原始数据")

            st.dataframe(total_df.head(), use_container_width=True)

            st.subheader("分表原始数据")

            st.dataframe(sub_df.head(), use_container_width=True)

       

        # 字段配置

        st.subheader("⚙️ 字段配置")

       

        # 定义常见字段名称用于自动匹配

        common_fields = {

            'commission': ['手续费', '佣金', '手续费用', 'commission', '费用'],

            'channel': ['渠道维护费', '渠道费', 'channel_fee', '维护费'],

            'operation': ['运营管理费', '运营费', 'operation_fee', '管理费'],

            'premium': ['保费', '保险费', '保费金额', '保单保费', 'premium', '保险金额'],

            'total_policy': ['分表单号', '子保单号', '分保单号', 'policy_main', '分单号'],

            'sub_policy': ['保单号', '保单编号', '子保单', 'policy_sub', '保单']

        }

       

        # 自动匹配字段的函数

        def find_default_field(df, candidates, optional=False):

            """根据候选名称自动匹配字段"""

            for col in df.columns:

                if col in candidates:

                    return col

            return None if optional else df.columns[0]  # 可选字段返回None，必要字段返回第一列

       

        col1, col2 = st.columns(2)

        with col1:

            st.write("**税率配置**")

            tax_rate = st.number_input("税率（默认1.06）", value=1.06, min_value=0.1, max_value=10.0, step=0.01)

        with col2:

            st.write("")

       

st.divider()
        st.write("**总表字段配置**")

        # 获取总表的默认值
        total_commission_default = find_default_field(total_df, common_fields['commission'])
        total_channel_default = find_default_field(total_df, common_fields['channel'], optional=True)
        total_operation_default = find_default_field(total_df, common_fields['operation'], optional=True)
        total_premium_default = find_default_field(total_df, common_fields['premium'])
        total_policy_default = find_default_field(total_df, common_fields['total_policy'])

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            total_commission_col = st.selectbox(
                "手续费字段",
                total_df.columns,
                index=list(total_df.columns).index(total_commission_default),
                key='total_commission'
            )
        with col2:
            channel_options = [None] + list(total_df.columns)
            total_channel_idx = channel_options.index(total_channel_default) if total_channel_default in channel_options else 0
            total_channel_col = st.selectbox(
                "渠道维护费字段",
                channel_options,
                format_func=lambda x: '无' if x is None else str(x), # 修复点1：使用更简洁的format_func
                index=total_channel_idx,
                key='total_channel'
            )
        with col3:
            op_options = [None] + list(total_df.columns)
            total_op_idx = op_options.index(total_operation_default) if total_operation_default in op_options else 0
            total_operation_col = st.selectbox(
                "运营管理费字段",
                op_options,
                format_func=lambda x: '无' if x is None else str(x), # 修复点1
                index=total_op_idx,
                key='total_operation'
            )
        with col4:
            total_premium_col = st.selectbox(
                "保费字段",
                total_df.columns,
                index=list(total_df.columns).index(total_premium_default),
                key='total_premium'
            )

        col1, col2 = st.columns(2)
        with col1:
            total_policy_col = st.selectbox(
                "分表单号字段",
                total_df.columns,
                index=list(total_df.columns).index(total_policy_default),
                key='total_policy'
            )

        st.divider()
        st.write("**分表字段配置**")

        # 获取分表的默认值
        sub_commission_default = find_default_field(sub_df, common_fields['commission'])
        sub_channel_default = find_default_field(sub_df, common_fields['channel'], optional=True)
        sub_operation_default = find_default_field(sub_df, common_fields['operation'], optional=True)
        sub_premium_default = find_default_field(sub_df, common_fields['premium'])
        sub_policy_default = find_default_field(sub_df, common_fields['sub_policy'])

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            sub_commission_col = st.selectbox(
                "手续费字段",
                sub_df.columns,
                index=list(sub_df.columns).index(sub_commission_default),
                key='sub_commission'
            )
        with col2:
            sub_channel_options = [None] + list(sub_df.columns)
            sub_channel_idx = sub_channel_options.index(sub_channel_default) if sub_channel_default in sub_channel_options else 0
            sub_channel_col = st.selectbox(
                "渠道维护费字段",
                sub_channel_options,
                format_func=lambda x: '无' if x is None else str(x), # 修复点1：防止与总表串位
                index=sub_channel_idx,
                key='sub_channel'
            )
        with col3:
            sub_op_options = [None] + list(sub_df.columns)
            sub_op_idx = sub_op_options.index(sub_operation_default) if sub_operation_default in sub_op_options else 0
            sub_operation_col = st.selectbox(
                "运营管理费字段",
                sub_op_options,
                format_func=lambda x: '无' if x is None else str(x), # 修复点1
                index=sub_op_idx,
                key='sub_operation'
            )
        with col4:
            sub_premium_col = st.selectbox(
                "保费字段",
                sub_df.columns,
                index=list(sub_df.columns).index(sub_premium_default),
                key='sub_premium'
            )

        col1, col2 = st.columns(2)
        with col1:
            sub_policy_col = st.selectbox(
                "保单号字段",
                sub_df.columns,
                index=list(sub_df.columns).index(sub_policy_default),
                key='sub_policy'
            )

        # 处理数据
        if st.button("🚀 执行处理", type='primary'):
            try:
                # 修复点2：增加专门的数据转换安全函数，抛出精确错误提示
                def to_float(df, col_name, table_desc, field_desc):
                    if not col_name:
                        return 0
                    try:
                        return df[col_name].astype(float)
                    except Exception as e:
                        # 主动抛出详细的错误信息
                        raise ValueError(f"【{table_desc}】的 {field_desc}（对应列名：'{col_name}'）包含无法转为数字的字符（如文本、空格等）。\n系统报错：{e}")

                total_df_copy = total_df.copy()
                sub_df_copy = sub_df.copy()

                # 使用安全函数计算总表
                total_commission = to_float(total_df_copy, total_commission_col, "总表", "手续费字段")
                total_channel = to_float(total_df_copy, total_channel_col, "总表", "渠道维护费字段")
                total_operation = to_float(total_df_copy, total_operation_col, "总表", "运营管理费字段")
                total_total_commission = total_commission + total_channel + total_operation
                total_net_premium = to_float(total_df_copy, total_premium_col, "总表", "保费字段") / tax_rate

                total_df_copy['费率'] = (total_total_commission / total_net_premium * 100).round(2)

                # 使用安全函数计算分表
                sub_commission = to_float(sub_df_copy, sub_commission_col, "分表", "手续费字段")
                sub_channel = to_float(sub_df_copy, sub_channel_col, "分表", "渠道维护费字段")
                sub_operation = to_float(sub_df_copy, sub_operation_col, "分表", "运营管理费字段")
                sub_total_commission = sub_commission + sub_channel + sub_operation
                sub_net_premium = to_float(sub_df_copy, sub_premium_col, "分表", "保费字段") / tax_rate

                sub_df_copy['分表费率'] = (sub_total_commission / sub_net_premium * 100).round(2)

                # （后续匹配代码与原代码一致，保持不变）...
                rate_mapping = dict(zip(
                    total_df_copy[total_policy_col].astype(str),
                    total_df_copy['费率']
                ))
                
                sub_df_copy['总表费率'] = sub_df_copy[sub_policy_col].astype(str).map(rate_mapping)
                sub_df_copy['费率差额'] = (sub_df_copy['总表费率'] - sub_df_copy['分表费率']).round(2)

                st.success("✅ 数据处理完成！")

               

                # 显示处理结果

                st.subheader("📈 处理结果")

        

                # 显示结果时添加百分号符号

                result_display = sub_df_copy.copy()

                if '分表费率' in result_display.columns:

                    result_display['分表费率'] = result_display['分表费率'].astype(str) + '%'

                if '总表费率' in result_display.columns:

                    result_display['总表费率'] = result_display['总表费率'].astype(str) + '%'

                if '费率差额' in result_display.columns:

                    result_display['费率差额'] = result_display['费率差额'].astype(str) + '%'

               

                st.dataframe(result_display, use_container_width=True)

               

                # 统计信息

                col1, col2, col3 = st.columns(3)

                with col1:

                    matched = sub_df_copy['总表费率'].notna().sum()

                    st.metric("匹配成功条数", matched)

                with col2:

                    unmatched = sub_df_copy['总表费率'].isna().sum()

                    st.metric("未匹配条数", unmatched)

                with col3:

                    total_records = len(sub_df_copy)

                    st.metric("总条数", total_records)

               

                # 费率统计

                st.subheader("📊 费率统计")

                col1, col2, col3, col4 = st.columns(4)

                with col1:

                    avg_sub_rate = sub_df_copy['分表费率'].mean()

                    st.metric("分表平均费率", f"{avg_sub_rate:.2f}%")

                with col2:

                    avg_total_rate = sub_df_copy['总表费率'].mean()

                    st.metric("总表平均费率", f"{avg_total_rate:.2f}%")

                with col3:

                    min_diff = sub_df_copy['费率差额'].min()

                    st.metric("最小费率差额", f"{min_diff:.2f}%")

                with col4:

                    max_diff = sub_df_copy['费率差额'].max()

                    st.metric("最大费率差额", f"{max_diff:.2f}%")

               

                # 下载结果

                st.subheader("📥 下载结果")

               

                # 转换为Excel

                output = io.BytesIO()

                with pd.ExcelWriter(output, engine='openpyxl') as writer:

                    sub_df_copy.to_excel(writer, sheet_name='结果', index=False)

                    total_df_copy.to_excel(writer, sheet_name='总表费率', index=False)

               

                output.seek(0)

               

                st.download_button(

                    label="📥 下载处理结果",

                    data=output.getvalue(),

                    file_name="excel_processing_result.xlsx",

                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                )

               
except ValueError as ve:
                # 捕获我们自定义的字段转换错误，展示给用户
                st.error(f"❌ 数据内容错误: {str(ve)}")
                st.info("💡 请检查对应的 Excel 列，确保没有混入文字（如姓名、车牌号）或其他非数字符号。")
            except Exception as e:
                st.error(f"❌ 处理出错: {str(e)}")
                st.info("请确保选择的字段存在且数据类型正确")



else:

    st.info("👆 请上传总表和分表Excel文件开始处理")
