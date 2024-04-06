import os 
from dotenv import load_dotenv
load_dotenv()

Api_key = os.getenv("API_KEY")

print(Api_key)