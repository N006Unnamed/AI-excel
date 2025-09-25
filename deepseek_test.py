from openai import OpenAI
import pandas as pd
from io import StringIO

client = OpenAI(api_key="sk-9fc773ee795847d1b71d5fd2ea3546fd", base_url="https://api.deepseek.com")

# path = "test2.xlsx"
# exl = pd.read_excel(path, sheet_name="Sheet1", engine='openpyxl',na_filter=False)  # sheet_name不指定时返回全表数据
# excel_text = exl.to_string(index=False)
#
# print(excel_text)

df = pd.read_excel('test2.xlsx', engine='openpyxl')
df.fillna(0, inplace=True)
print(df)
excel_text = df.to_string(index=False)

response = client.chat.completions.create(
    model="deepseek-chat",
    messages=[
        {"role": "system", "content": "You are a helpful assistant"},
        {"role": "user", "content": "请填充excel表中的数据，其中第二行“成本估算”的值等于其他行费用项的值相加的值，"
                                    "例如2026年成本估算的值为："
                                    "660+510+500+74.71+300+300+200+0+87.3+297.14+239.94+483.6+270+0+200=4122.70。"
                                    "第三列“合计”为单项费用所有年份列的值相加，例如人工成本的合计为："
                                    "660+660+660+673.2+673.2+673.2+686.66+686.66+686.66+700.40+700.40+700.40+714.40+714.40=9589.59。"
                                    "其中最后一列不参与计算。注意，输出的结果中，第一列与第二列、第二列与第三列之间要用3个空格符间隔，"
                                    "确保计算结果准确，不要自己捏造数据。"
                                    "并且每一行中项与项之间用3个空格符作为间隔，例如输出“二   成本估算   4122.69    4122.69”。"
                                    "仅输出填充完成后的excel结果，包括表头，除此之外不要输出其他文字"+excel_text},
    ],
    stream=False
)


print(response.choices[0].message.content)

data = response.choices[0].message.content

# 使用StringIO将字符串转换为类似文件的对象
data_io = StringIO(data.strip())

# 读取数据到DataFrame
# 设置分隔符为多个空格，跳过初始空格，并处理不规则列名
df = pd.read_csv(
    data_io,
    sep=r'\s{2,}',  # 匹配2个或更多连续空格作为分隔符
    engine='python',
    skipinitialspace=True,
    header=0,
    dtype=str  # 先读取为字符串，后续再转换类型
)

# 重命名列
df.columns = ['  ', '  ', '合计'] + [f'{i}' for i in range(2026, 2040)] + [' ']

# 保存为Excel文件
excel_file = '成本估算数据.xlsx'
df.to_excel(excel_file, index=False)

print(f"Excel文件已成功保存到: {excel_file}")

# reasoning_content = ""
# content = ""
#
# for chunk in response:
#     if chunk.choices[0].delta.reasoning_content:
#         reasoning_content += chunk.choices[0].delta.reasoning_content
#     else:
#         content += chunk.choices[0].delta.content




# from openai import OpenAI
#
# # 初始化 OpenAI 客户端
# client = OpenAI(api_key="sk-9fc773ee795847d1b71d5fd2ea3546fd", base_url="https://api.deepseek.com")
#
# # 读取本地 text.txt 文件
# with open("./file.txt", "r", encoding="utf-8") as file:
#     file_content = file.read()
#
# # 将文件内容发送给大模型
# response = client.chat.completions.create(
#     model="deepseek-chat",
#     messages=[
#         {"role": "system", "content": "You are a helpful assistant"},
#         {"role": "user", "content": "帮我总结下面文档中的内容"+file_content},
#     ],
#     stream=False
# )
#
# # 打印大模型的响应
# print(response.choices[0].message.content)

# from openai import OpenAI
# import pandas as pd
#
# # 初始化 OpenAI 客户端
# client = OpenAI(api_key="换成自己的key", base_url="https://api.deepseek.com")
#
# # 读取本地 text.txt 文件
# with open("./file.txt", "r", encoding="utf-8") as file:
#     file_content = file.read()
#
# # 将文件内容发送给大模型
# response = client.chat.completions.create(
#     model="deepseek-chat",
#     messages=[
#         {"role": "system", "content": "You are a helpful assistant"},
#         {"role": "user", "content": "帮我总结下面文档中的内容"+file_content},
#     ],
#     stream=False
# )
#
# # 打印大模型的响应
# print(response.choices[0].message.content)