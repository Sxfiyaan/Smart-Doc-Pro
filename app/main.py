from langchain_ollama import ChatOllama

from app.core.config import OLLAMA_BASE_URL, OLLAMA_MODEL


print("Ollama URL   :", OLLAMA_BASE_URL)
print("Ollama Model :", OLLAMA_MODEL)


llm = ChatOllama(
    model=OLLAMA_MODEL,
    base_url=OLLAMA_BASE_URL,
    temperature=0
)

response = llm.invoke(
    "Explain what an Excel spreadsheet is in one sentence."
)

print("\nModel response:")
print(response.content)