"""Test LLM configuration and availability"""

import os
from dotenv import load_dotenv
from llm_integration import get_llm_formatter

# Load environment variables from .env file
load_dotenv()

def test_llm_config():
    """Test if LLM is properly configured and available"""
    
    print("🔍 Checking LLM Configuration...")
    print("-" * 50)
    
    # Check environment variables
    openai_key = os.getenv('OPENAI_API_KEY')
    anthropic_key = os.getenv('ANTHROPIC_API_KEY')
    azure_key = os.getenv('AZURE_OPENAI_API_KEY')
    
    print("✅ Environment Variables Found:")
    if openai_key:
        print(f"  - OPENAI_API_KEY: {'*' * 10}{openai_key[-4:] if len(openai_key) > 4 else '****'}")
        print(f"  - OPENAI_MODEL: {os.getenv('OPENAI_MODEL', 'gpt-4o-mini (default)')}")
    
    if anthropic_key:
        print(f"  - ANTHROPIC_API_KEY: {'*' * 10}{anthropic_key[-4:] if len(anthropic_key) > 4 else '****'}")
        print(f"  - ANTHROPIC_MODEL: {os.getenv('ANTHROPIC_MODEL', 'claude-3-haiku-20240307 (default)')}")
    
    if azure_key:
        print(f"  - AZURE_OPENAI_API_KEY: {'*' * 10}{azure_key[-4:] if len(azure_key) > 4 else '****'}")
        print(f"  - AZURE_OPENAI_ENDPOINT: {os.getenv('AZURE_OPENAI_ENDPOINT', 'Not set')}")
        print(f"  - AZURE_OPENAI_DEPLOYMENT: {os.getenv('AZURE_OPENAI_DEPLOYMENT', 'Not set')}")
    
    if not (openai_key or anthropic_key or azure_key):
        print("  ❌ No LLM API keys found in environment")
    
    print("\n📦 Testing LLM Formatter Initialization...")
    print("-" * 50)
    
    # Test LLM formatter
    formatter = get_llm_formatter()
    
    if formatter:
        print(f"✅ LLM Formatter Active!")
        print(f"  - Provider: {formatter.provider}")
        print(f"  - Model: {formatter.model}")
        print(f"  - Ready for text formatting: Yes")
        
        # Test formatting
        print("\n🧪 Testing Text Formatting...")
        print("-" * 50)
        
        test_text = "Implementing new security protocols and conducting vulnerability assessments. Working on system integration with third-party APIs."
        formatted = formatter.format_text(test_text, "activities")
        
        print(f"Original text: {test_text}")
        print(f"\nFormatted text: {formatted}")
        
        # Test placeholder detection
        test_placeholder = "Owner to share risks"
        formatted_placeholder = formatter.format_text(test_placeholder, "risks")
        print(f"\nPlaceholder test:")
        print(f"  Input: '{test_placeholder}'")
        print(f"  Output: '{formatted_placeholder}'")
        
    else:
        print("❌ LLM Formatter not available")
        print("  - Text will be used as-is without formatting")
        print("\n💡 To enable LLM formatting:")
        print("  1. Make sure you have installed: pip install openai anthropic")
        print("  2. Set API keys in .env file")
        print("  3. Restart the Flask application")

if __name__ == "__main__":
    test_llm_config()
