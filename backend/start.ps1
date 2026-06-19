# Set your Anthropic API key here
$env:ANTHROPIC_API_KEY = "YOUR_API_KEY_HERE"

# Start the FastAPI backend
$uvicornPath = "$env:APPDATA\Python\Python314\Scripts\uvicorn.exe"
& $uvicornPath main:app --reload --port 8000
