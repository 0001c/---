import os
from openai import OpenAI

models = [
    "doubao-seed-2-0-lite-260215",
    "doubao-seed-2-0-mini-260215",
    "glm-4-7-251222"
]

def get_ai_response(client, model, prompt):
    """封装大模型调用和解析的通用函数"""
    tools = [{"type": "web_search", "max_keyword": 3, "limit": 6}]
    target_text = "未获取到分析结果，请稍后重试。" # 设置默认值
    
    try:
        response = client.responses.create(
            model=model,
            tools=tools,
            input=[{"role": "user", "content": [{"type": "input_text", "text": prompt}]}]
        )
        
        # 解析逻辑
        output_messages = [item for item in response.output if item.type == 'message']
        if output_messages:
            output_texts = [content for content in output_messages[0].content if content.type == 'output_text']
            if output_texts:
                target_text = output_texts[0].text
                print("提取成功！")
            else:
                print("未找到output_text类型的内容")
        else:
            print("未找到message类型的output项")
            
    except Exception as e:
        print(f"大模型请求或提取失败，错误信息：{e}")
        
    return target_text

def main(data):
    print("开始处理数据")
    result = []
    api_key = os.getenv('ARK_API_KEY')
    client = OpenAI(
        base_url="https://ark.cn-beijing.volces.com/api/v3",
        api_key=api_key,
    )

    # 1. 分析学员调剂可能性
    print("分析学员调剂可能性...")
    prompt_1 = f"以下是某位考研考生考研调剂意向调查的基本信息，请根据这些信息，你也可以网上搜索相关数据，对该考生学员调剂可能性分析。\n{data},要求回复内容严格按以下格式，不用markdown格式，不要凭空假设猜想，要求回复内容要相关、准确、详细、确定有权威依据。\n1.分数整体分析：\n2.本科院校及专业：\n3.一志愿院校：\n4.语种调剂难度：\n5.作品集等成果：\n6.学习时间及精力：\n....."
    result.append(get_ai_response(client, models[0], prompt_1))

    # 2. 推荐院校
    print("推荐院校...")
    prompt_2 = f'''以下是某位考研考生考研调剂意向调查的基本信息，请根据这些信息，你也可以网上搜索相关数据确保数据时效性，今年是2026年，对该考生给出往年调剂大数据下三档院校参考。\n{data},要求回复内容严格按以下格式，不用markdown格式，不要凭空假设猜想，要求回复内容要相关、准确、详细、确定有权威依据。\n
a区推荐院校
冲：
稳：
保：
b区推荐院校
冲：
稳：
保：

建议：
                                '''
    result.append(get_ai_response(client, models[1], prompt_2))
   
    return result
