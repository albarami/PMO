"""Optional LLM Integration for Text Formatting Only"""

import os
import logging
from typing import Optional, Dict, Any

# LLM libraries are optional
try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False


logger = logging.getLogger(__name__)


class LLMFormatter:
    """LLM text formatter for PMO reports - NO calculations allowed."""
    
    SYSTEM_PROMPT = """You are a PMO report assistant. Your role is to help format and clarify project status information.

CRITICAL RULES:
1. NEVER calculate or estimate any numbers - all metrics are provided
2. NEVER make up data - only reformat/clarify what's provided
3. Keep responses concise and professional
4. If text is unclear or placeholder (like "Owner to share"), return "[To be provided]"
5. Format text as clean bullet points when appropriate
6. Maximum 3-4 bullet points for activities
7. Keep each bullet point to one line"""
    
    def __init__(self):
        """Initialize LLM formatter with available providers."""
        self.provider = None
        self.client = None
        self.model = None
        
        # Try to initialize available LLM providers
        # Try OpenAI first, then Anthropic, then Azure
        self._init_openai()
        if not self.client:
            self._init_anthropic()
        if not self.client:
            self._init_azure_openai()
    
    def _init_openai(self):
        """Initialize OpenAI if available and configured."""
        if not OPENAI_AVAILABLE:
            return
        
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            return
        
        try:
            self.client = openai.OpenAI(api_key=api_key)
            self.provider = 'openai'
            self.model = os.getenv('OPENAI_MODEL', 'gpt-4o-mini')
            logger.info(f"Initialized OpenAI with model {self.model}")
        except Exception as e:
            logger.error(f"Failed to initialize OpenAI: {e}")
    
    def _init_anthropic(self):
        """Initialize Anthropic if available and configured."""
        if not ANTHROPIC_AVAILABLE:
            return
        
        api_key = os.getenv('ANTHROPIC_API_KEY')
        if not api_key:
            return
        
        try:
            self.client = anthropic.Anthropic(api_key=api_key)
            self.provider = 'anthropic'
            self.model = os.getenv('ANTHROPIC_MODEL', 'claude-3-haiku-20240307')
            logger.info(f"Initialized Anthropic with model {self.model}")
        except Exception as e:
            logger.error(f"Failed to initialize Anthropic: {e}")
    
    def _init_azure_openai(self):
        """Initialize Azure OpenAI if available and configured."""
        if not OPENAI_AVAILABLE:
            return
        
        api_key = os.getenv('AZURE_OPENAI_API_KEY')
        endpoint = os.getenv('AZURE_OPENAI_ENDPOINT')
        deployment = os.getenv('AZURE_OPENAI_DEPLOYMENT')
        
        if not all([api_key, endpoint, deployment]):
            return
        
        try:
            self.client = openai.AzureOpenAI(
                api_key=api_key,
                api_version="2023-12-01-preview",
                azure_endpoint=endpoint
            )
            self.provider = 'azure'
            self.model = deployment
            logger.info(f"Initialized Azure OpenAI with deployment {self.model}")
        except Exception as e:
            logger.error(f"Failed to initialize Azure OpenAI: {e}")
    
    def is_available(self) -> bool:
        """Check if LLM is available for text formatting."""
        return self.client is not None
    
    def format_text(self, text: str, context: str = "activities") -> str:
        """
        Format text using LLM if available, otherwise return original.
        
        Args:
            text: The text to format
            context: What type of text (activities, risks, issues)
        
        Returns:
            Formatted text or original if LLM not available
        """
        if not self.client or not text or text.strip() == '':
            return text
        
        # Check for placeholder text
        placeholder_phrases = [
            'owner to share',
            'to be provided',
            'tbd',
            'n/a',
            'na',
            'none',
            'nil'
        ]
        
        if any(phrase in text.lower() for phrase in placeholder_phrases):
            return '[To be provided]'
        
        try:
            if self.provider == 'openai' or self.provider == 'azure':
                return self._format_with_openai(text, context)
            elif self.provider == 'anthropic':
                return self._format_with_anthropic(text, context)
        except Exception as e:
            logger.error(f"LLM formatting failed: {e}")
            return text
        
        return text
    
    def _format_with_openai(self, text: str, context: str) -> str:
        """Format text using OpenAI."""
        prompt = self._get_prompt(text, context)
        
        try:
            if self.provider == 'azure':
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": self.SYSTEM_PROMPT},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3,
                    max_tokens=200
                )
            else:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": self.SYSTEM_PROMPT},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.3,
                    max_tokens=200
                )
            
            return response.choices[0].message.content.strip()
        except Exception as e:
            logger.error(f"OpenAI formatting error: {e}")
            return text
    
    def _format_with_anthropic(self, text: str, context: str) -> str:
        """Format text using Anthropic."""
        prompt = self._get_prompt(text, context)
        
        try:
            response = self.client.messages.create(
                model=self.model,
                system=self.SYSTEM_PROMPT,
                messages=[
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=200
            )
            
            return response.content[0].text.strip()
        except Exception as e:
            logger.error(f"Anthropic formatting error: {e}")
            return text
    
    def _get_prompt(self, text: str, context: str) -> str:
        """Generate prompt based on context."""
        if context == "activities":
            return f"""Format the following project activities into clear bullet points (max 3-4 points).
Keep each point concise and professional.

Text: {text}

Formatted output:"""
        
        elif context == "risks":
            return f"""Clarify and format the following project risks into a brief, clear description.
Keep it concise and professional.

Text: {text}

Formatted output:"""
        
        elif context == "issues":
            return f"""Clarify and format the following project issues into a brief, clear description.
Keep it concise and professional.

Text: {text}

Formatted output:"""
        
        else:
            return f"""Format the following text to be clear and professional.
Do not add any information not present in the original.

Text: {text}

Formatted output:"""
    
    def format_activities_batch(self, activities_list: list) -> list:
        """Format multiple activity texts efficiently."""
        if not self.is_available():
            return activities_list
        
        formatted = []
        for activity in activities_list:
            formatted.append(self.format_text(activity, "activities"))
        
        return formatted


# Global formatter instance
llm_formatter = None


def get_llm_formatter() -> Optional[LLMFormatter]:
    """Get or create the global LLM formatter instance."""
    global llm_formatter
    if llm_formatter is None:
        llm_formatter = LLMFormatter()
    return llm_formatter if llm_formatter.is_available() else None


def format_project_text(project_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Format text fields in project data using LLM if available.
    
    This function ONLY formats text - it NEVER modifies numbers or calculations.
    """
    formatter = get_llm_formatter()
    if not formatter:
        return project_data
    
    # Create a copy to avoid modifying original
    formatted_data = project_data.copy()
    
    # Format text fields only
    text_fields = {
        'current_activities': 'activities',
        'future_activities': 'activities',
        'risks': 'risks',
        'issues': 'issues',
        'comments': 'general'
    }
    
    for field, context in text_fields.items():
        if field in formatted_data and formatted_data[field]:
            original = formatted_data[field]
            # Only format if it's not already a placeholder
            if not original.startswith('[') and original != 'TBD':
                formatted = formatter.format_text(original, context)
                formatted_data[field] = formatted
    
    return formatted_data
