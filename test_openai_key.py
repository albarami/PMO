"""Test OpenAI API Key directly"""

import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def test_openai_key():
    """Test if OpenAI key works"""
    print("🔑 Testing OpenAI API Key...")
    print("=" * 60)
    
    # Get the key
    api_key = os.getenv('OPENAI_API_KEY')
    
    if not api_key:
        print("❌ No OPENAI_API_KEY found in .env file")
        return
    
    # Show key info (partial for security)
    print(f"✅ Key found: {api_key[:8]}...{api_key[-4:]}")
    print(f"   Key length: {len(api_key)} characters")
    
    # Check key format
    if api_key.startswith('sk-'):
        print("✅ Key format looks correct (starts with 'sk-')")
    else:
        print("⚠️ Key might have wrong format (should start with 'sk-')")
    
    # Try to use the key
    try:
        from openai import OpenAI
        
        print("\n🧪 Testing API connection...")
        client = OpenAI(api_key=api_key)
        
        # Make a minimal test request
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",  # Using cheaper model for test
            messages=[{"role": "user", "content": "Say 'API working'"}],
            max_tokens=10
        )
        
        result = response.choices[0].message.content
        print(f"✅ API Response: {result}")
        print("\n🎉 OpenAI API key is VALID and WORKING!")
        
    except Exception as e:
        print(f"\n❌ API Error: {str(e)}")
        
        # Check for common issues
        if "401" in str(e):
            print("\n💡 Possible issues:")
            print("  1. Key might be expired or revoked")
            print("  2. Key might have extra spaces or quotes")
            print("  3. Key might be for wrong OpenAI account/project")
            print("\n📝 To fix:")
            print("  1. Get a new key from: https://platform.openai.com/api-keys")
            print("  2. Make sure to copy EXACTLY without spaces")
            print("  3. Update .env file with: OPENAI_API_KEY=sk-...")
        elif "429" in str(e):
            print("\n⚠️ Rate limit or quota exceeded")
            print("  - Check your OpenAI account credits")
            print("  - You might need to add payment method")

if __name__ == "__main__":
    test_openai_key()
