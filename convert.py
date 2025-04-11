from openai import OpenAI

client = OpenAI(api_key="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")  # Or rely on environment variable

response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "user", "content": "can i use chatgpt for UI automation?"}
    ]
)

print(response.choices[0].message.content)



 
