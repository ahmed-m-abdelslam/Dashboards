import os

from groq import Groq # type: ignore

groq_api = os.getenv("GROQ_API")

client = Groq(api_key=groq_api)
completion = client.chat.completions.create(
    model="openai/gpt-oss-safeguard-20b",
    messages=[
      {
        "role": "user",
        "content": """
        summary this : {Data}
        """
      }
    ],
    temperature=1,
    max_completion_tokens=8192,
    top_p=1,
    reasoning_effort="medium",
    stream=True,
    stop=None
)

for chunk in completion:
    print(chunk.choices[0].delta.content or "", end="")
