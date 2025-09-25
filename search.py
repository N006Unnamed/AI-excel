import requests
import os
import re
import json
from datetime import datetime
from openpyxl_vba import load_workbook
import sys

# DeepSeek API 配置
DEEPSEEK_API_KEY = "sk-9fc773ee795847d1b71d5fd2ea3546fd"  # 替换为你的 DeepSeek API 密钥
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
MODEL_NAME = "deepseek-chat"  # 使用 DeepSeek 的 chat 模型


def read_financial_params(input_file, sheet_name):
    """从Excel文件读取财务参数数据"""
    print(f"正在读取Excel文件: {input_file}")

    try:
        # 打开工作簿
        wb = load_workbook(input_file)
        sheet = wb[sheet_name]

        if sheet_name == "A财务假设":
            # 确定表头位置
            header_row = 2

            # 识别列索引
            headers = {}
            headers["id"] = 1
            headers["name"] = 2
            headers["source"] = 3

            # 检查必要的列
            if "id" not in headers or "name" not in headers or "source" not in headers:
                raise ValueError("Excel文件中缺少必要的列: 序号、名称、解释/参考")

            # 提取参数数据
            params = []
            for row in sheet.iter_rows(min_row=header_row + 1):
                # 跳过空行
                if not row[headers["id"] - 1].value:
                    continue

                param = {
                    "id": row[headers["id"] - 1].value,
                    "name": row[headers["name"] - 1].value,
                    "source": row[headers["source"] - 1].value if row[headers["source"] - 1].value else ""
                }
                params.append(param)

            print(f"成功读取 {len(params)} 个财务参数")
            return params

        else:
            # 确定表头位置
            header_row = 2

            # 识别列索引
            headers = {}
            headers["id"] = 1
            headers["name"] = 2

            # 检查必要的列
            if "id" not in headers or "name" not in headers :
                raise ValueError("Excel文件中缺少必要的列: 序号、名称")

            # 提取参数数据
            params = []
            for row in sheet.iter_rows(min_row=header_row + 1):
                # 跳过空行
                if not row[headers["id"] - 1].value:
                    continue

                param = {
                    "id": row[headers["id"] - 1].value,
                    "name": row[headers["name"] - 1].value
                }
                params.append(param)

            print(f"成功读取 {len(params)} 个财务参数")
            return params

    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        sys.exit(1)


def get_param_value(param_name, source_info, cache=True):
    """使用 DeepSeek API 获取财务参数的具体数值（非范围区间）"""
    # 创建缓存目录
    if not os.path.exists("cache"):
        os.makedirs("cache")
    # 检查缓存
    cache_filename = f"cache/{param_name.replace(' ', '_').replace('/', '_')}.json"
    if cache and os.path.exists(cache_filename):
        with open(cache_filename, "r", encoding="utf-8") as f:
            cached_data = json.load(f)
            print(f"使用缓存值: {param_name} -> {cached_data['value']}")
            return cached_data["value"]

    # 构建专业提示词 - 强调需要具体数值而非范围
    prompt = f"""
    ## 任务说明
    作为财务专家，请根据中国最新政策和标准，提供以下财务参数的具体数值（不要范围区间）：

    ## 参数信息
    参数名称: {param_name}
    背景信息: {source_info}

    ## 要求
    1. 必须提供具体数值（如5%、8%等），不要给出范围（如8-12%）
    2. 如果是百分比，请以%表示
    3. 如果因行业/地区不同而有差异，请选择最常用或标准值
    4. 回答尽量简洁，不超过20字
    5. 数值必须基于中国最新的财务政策和标准
    6. 如果背景信息不为空则参考背景信息中的信息，如果为空则无视
    7. 无标准税率时默认使用5%

    ## 示例格式
    项目投资财务基准收益率(税前): 10%
    城市维护建设税税率(市区): 7%
    教育费附加费率: 3%
    """

    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system",
             "content": "你是一位资深财务分析师，熟悉中国财务政策和标准。请严格按照要求提供具体数值，不要给出范围区间。"},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.1,  # 低随机性确保结果准确
        "max_tokens": 100
    }

    try:
        print(f"查询参数: {param_name}...")
        response = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload)
        response.raise_for_status()

        # 解析响应
        result = response.json()
        if "choices" not in result or len(result["choices"]) == 0:
            raise ValueError("API响应格式错误")

        value = result["choices"][0]["message"]["content"].strip()

        # 清理结果：移除多余描述，提取核心数值
        if "：" in value:
            value = value.split("：")[-1].strip()
        if ":" in value:
            value = value.split(":")[-1].strip()

        # 确保数值格式规范 - 移除范围指示词
        value = re.sub(r'约|大约|通常|一般|左右|之间|不等|约为|大约为|至|到|~|-', '', value).strip()

        # 处理百分比格式
        if "百分之" in value:
            value = value.replace("百分之", "") + "%"
        if "%" not in value and re.search(r'[0-9]+\.?[0-9]*$', value):
            if "率" in param_name or "比例" in param_name or "费率" in param_name:
                value += "%"

        # 保存到缓存
        if cache:
            cache_data = {
                "param_name": param_name,
                "source_info": source_info if source_info else "",
                "value": value,
                "timestamp": datetime.now().isoformat()
            }
            with open(cache_filename, "w", encoding="utf-8") as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)

        return value

    except requests.exceptions.RequestException as e:
        print(f"API请求失败: {e}")
        return "API请求失败"
    except Exception as e:
        print(f"处理API响应失败: {e}")
        return "解析失败"


def deepseek_tax(input_file):
    """
    使用 deepseek 查询 表A、G.2、G.3 税率
    :param input_file:
    :return:
    """

    wb = load_workbook(input_file)
    sheet_names = ["A财务假设", "G.2进项增值税率-运营成本", "G.3销项增值税率"]

    for sheet_name in sheet_names:
        ws = wb[sheet_name]

        params = read_financial_params(input_file, sheet_name)

        if sheet_name == "A财务假设":
            # 填充数据
            for row_num, param in enumerate(params, start=3):
                # 获取参数值
                param_value = get_param_value(param["name"], param["source"])

                # 写入参数值
                ws.cell(row=row_num, column=4).value = param_value
        else:
            # 填充数据
            for row_num, param in enumerate(params, start=3):
                # 获取参数值
                param_value = get_param_value(param["name"], "")

                # 写入参数值
                ws.cell(row=row_num, column=3).value = param_value

    # 保存文件
    wb.save(input_file)
    print(f"结果Excel文件已保存: {input_file}")

def get_min_turnover_days(api_key):
    """
    通过 DeepSeek API 获取各项目的最低周转天数，适用于项目财务分析

    参数:
    api_key (str): DeepSeek API 密钥

    返回:
    str: 包含13个项目最低周转天数的字符串，格式为"1,2,3,4,5,6,7,8,9,10,11,12,13"
    """
    # API端点
    url = "https://api.deepseek.com/v1/chat/completions"

    # 请求头
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    # 更专业的财务分析提示词
    prompt = """作为财务分析专家，请为以下项目提供合理的最低周转天数估算（基于典型制造业项目）：
    1. 流动资产
    2. 应收账款
    3. 存货
    4. 原材料
    5. 燃料动力
    6. 在产品
    7. 产成品
    8. 现金
    9. 流动负债
    10. 应付账款
    11. 流动资金(流动资产-流动负债)
    12. 流动资金当年增加额
    13. 流动资金贷款

    请基于以下行业标准提供合理估算：
    - 周转天数都应该大于0
    - 周转天数都应该能被360整除
    - 应收账款周转天数通常为30-60天
    - 存货周转天数通常为45-90天
    - 原材料周转天数通常为15-30天
    - 在产品周转天数通常为10-20天
    - 产成品周转天数通常为15-30天
    - 应付账款周转天数通常为30-45天

    请只返回13个数字，用英文逗号分隔，不要包含任何其他文本。"""

    # 请求体
    payload = {
        "model": "deepseek-chat",
        "messages": [
            {
                "role": "system",
                "content": "你是一个专业的财务分析师，专注于项目投资评估和财务分析。请提供基于行业标准的合理估算。"
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0.3  # 适度创造性，但保持专业性
    }

    try:
        # 发送请求
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()  # 检查请求是否成功

        # 解析响应
        data = response.json()
        answer = data["choices"][0]["message"]["content"].strip()

        # 清理响应，确保只包含数字和逗号
        # 移除非数字和逗号字符
        cleaned_answer = ''.join(c for c in answer if c.isdigit() or c == ',')

        # 验证格式是否正确（应包含12个逗号和13个数字段）
        parts = cleaned_answer.split(',')
        if len(parts) != 13:
            print(parts)
            raise ValueError(f"返回的数字数量不正确，期望13个，得到{len(parts)}个")

        # 验证每个部分都是数字
        for part in parts:
            if not part.isdigit():
                raise ValueError(f"返回的内容包含非数字值: {part}")

        return cleaned_answer

    except requests.exceptions.RequestException as e:
        print(f"API请求错误: {e}")
        return None
    except (KeyError, IndexError) as e:
        print(f"API响应解析错误: {e}")
        return None
    except ValueError as e:
        print(f"响应格式错误: {e}")
        return None


def deepseek(input_file):
    # 替换为您的DeepSeek API密钥
    API_KEY = "sk-9fc773ee795847d1b71d5fd2ea3546fd"
    sheet_name = "D.4流动资金估算表"

    result = get_min_turnover_days(API_KEY)
    if result:
        print(f"最低周转天数: {result}")
        # 将结果拆分为列表以便进一步处理
        days_list = result.split(',')
        print("各项目周转天数:")
        items = [
            "流动资产", "应收账款", "存货", "原材料", "燃料动力", "在产品",
            "产成品", "现金", "流动负债", "应付账款", "流动资金(1-2)",
            "流动资金当年增加额", "流动资金贷款"
        ]
        wb = load_workbook(input_file)
        ws = wb[sheet_name]

        for i, (item, days) in enumerate(zip(items, days_list), 1):
            ws.cell(row=2 + i, column=3).value = days
            print(f"{i}. {item}: {days}天")

        wb.save(input_file)


    else:
        print("获取数据失败")


def ai_agent():
    api_key = "app-NYoSQ87p8dMcAMzNTe7RiM3Y"
    # 请求头
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    url = "http://172.16.0.115/v1"

    data = {
        "inputs": "你好，请介绍一下你自己",
        "response_mode": "streaming",
        "user": "441780824"
    }

    # 发送请求
    response = requests.post(url, headers=headers, data=json.dumps(data))

    if response.status_code == 200:
        print("状态码：", response.status_code)
        print("原始响应：", response.text[:500])  # 只打前500字符，防止太长
    else:
        print("调用失败：", response.status_code, response.text)
    response.raise_for_status()  # 检查请求是否成功
    print(response.text)


if __name__ == "__main__":
    # ai_agent()
    input_file = "财务分析套表自做模板编程用2.xlsm"
    deepseek(input_file)
    deepseek_tax(input_file)
