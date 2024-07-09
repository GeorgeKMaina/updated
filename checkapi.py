import os
from dotenv import load_dotenv
import openai
            
# Load environment variables from the .env file
load_dotenv()

api_key = os.getenv('new_api_key')

#if api_key is not None:
#    print(f"API Key: {api_key}")
#else:
#    print("API Key not found in .env file")

#print('hello',os.getenv('new_api_key'))
a=os.getenv('new_api_key')
print(api_key)


