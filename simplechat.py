import os
import openai
openai.api_key=os.getenv("new_api_key")

openai.api_key   # Set the API key directly

# Create a chat completion request
response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",  # Choose an appropriate model
    messages=[
        {"role": "user", "content": "Write for me an email stating i will be away for a while since i will be doing my exams"}
    ],
    max_tokens=150,
    n=1,
    stop=None,
    temperature=0.7,
)

# Extract and print the response content
print(response.choices[0].message['content'].strip())