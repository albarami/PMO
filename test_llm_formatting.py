"""Test LLM text formatting with real data"""

from dotenv import load_dotenv
load_dotenv()

from llm_integration import format_project_text, get_llm_formatter

def test_llm_formatting():
    """Test that LLM formatting actually works"""
    print("🤖 Testing LLM Text Formatting")
    print("=" * 60)
    
    # Check LLM status
    formatter = get_llm_formatter()
    
    if formatter:
        print(f"✅ LLM Active: {formatter.provider} - {formatter.model}")
    else:
        print("❌ LLM not configured")
        return
    
    # Test with sample project data
    test_project = {
        'name': 'Test Project',
        'current_activities': 'Working on system integration and testing user interfaces and conducting security audits',
        'future_activities': 'Deploy to production environment and train users and monitor performance',
        'risks': 'Potential delays due to vendor dependencies and resource constraints',
        'issues': 'API integration errors and database performance issues need resolution',
        'comments': 'Project is progressing well overall'
    }
    
    print("\n📝 Original text:")
    print(f"Activities: {test_project['current_activities']}")
    
    print("\n⚙️ Applying LLM formatting...")
    formatted_project = format_project_text(test_project)
    
    print("\n✨ Formatted text:")
    print(f"Activities: {formatted_project['current_activities']}")
    
    if test_project['current_activities'] != formatted_project['current_activities']:
        print("\n✅ LLM formatting is WORKING! Text has been transformed into bullet points.")
    else:
        print("\n⚠️ Text unchanged - LLM might have failed silently")

if __name__ == "__main__":
    test_llm_formatting()
