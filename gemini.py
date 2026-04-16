import google.generativeai as genai
import json
import re
import time
from google.api_core import exceptions

class GeminiAI:
    """Handle Gemini AI operations for email response parsing"""
    
    def __init__(self, api_key):
        self.api_key = api_key
        genai.configure(api_key=api_key)
        # Priority list of models to try
        self.models_to_try = [
            'gemini-2.0-flash',
            'gemini-1.5-pro',
            'gemini-1.5-flash',
            'gemini-flash-latest',
            'gemini-pro',
            'gemini-1.0-pro'
        ]
        self.model_name = self.models_to_try[0] # Default Start
    

    def parse_email_response(self, email_body, user_list):
        """
        Parse natural language email response using AI and context
        """
        # Create context string
        context_str = json.dumps([
            {'name': u['user_name'], 'email': u['email'], 'role': u.get('roles', '')} 
            for u in user_list
        ], indent=2)

        prompt = f"""
You are an intelligent email parser for a User Review System.
Your goal is to extract user management actions from the email body.

VALID USERS LIST (Context):
{context_str}

EMAIL BODY:
{email_body}

INSTRUCTIONS:
1. **Analyze** the email to find actions like "Delete [Name]", "Change [Name] role to [Role]", "Remove [Name] from [Role]", "Keep all", etc.
2. **Identify Users**: 
   - You MUST map names in the email (e.g. "Gaurav") to the EXACT `email` in the VALID USERS LIST.
   - If the user provides an Email (e.g. "gaurav@example.com"), verify it exists in the list.
   - **CRITICAL**: If a user is NOT in the VALID USERS LIST, DO NOT generate an action for them. Ignore them.
3. **Action Types**:
   - `delete`: For "Delete", "Remove", "Terminate".
   - `update_role`: For "Update role", "Change to", "Make [Role]", "Remove from [Role]".
   - `no_action`: For "Keep all", "No changes".

4. **Handling "Remove from [Role]"**:
   - If the instruction is "Remove [User] from [Role]", interpretation depends on intent. 
   - If they mean "remove the user entirely", use `delete`.
   - If they mean "remove that specific role" (but keep user), use `update_role` and set `new_value` to a role list WITHOUT that role.
   - If ambiguous, prefer `delete` only if the word "Delete" or "Remove user" is explicitly used. "Remove from Admin" usually means `update_role`.

OUTPUT FORMAT (JSON ONLY):
{{
    "actions": [
        {{
            "action_type": "delete|update_role",
            "user_email": "exact_email_from_list",
            "user_name": "name_from_list",
            "new_value": "new_role_if_update",
            "description": "Professional summary of the action"
        }}
    ]
}}
IMPORTANT: Return ONLY the raw JSON string. Do not use Markdown code blocks (```json).
"""
        
        # MODEL FALLBACK LOOP
        result_text = None
        last_error = None
        
        for model_name in self.models_to_try:
            try:
                model = genai.GenerativeModel(model_name)
                # Attempt generation
                response = model.generate_content(prompt)
                result_text = response.text.strip()
                # print(f"DEBUG: Success with model {model_name}")
                break # Success
            except exceptions.ResourceExhausted as e:
                print(f"WARNING: Quota exceeded for {model_name}. Trying next...")
                last_error = e
                continue
            except Exception as e:
                print(f"WARNING: Error with {model_name}: {e}. Trying next...")
                last_error = e
                continue

        if not result_text:
             print(f"Error parsing email after trying all models. Last error: {last_error}")
             return {"actions": [], "error": f"AI Error: {str(last_error)}"}

        # Extract JSON (Outside Loop)
        result_text = self._extract_json(result_text)
        print(f"DEBUG: EXTRACTED JSON: {result_text}")
            
        try:
            result = json.loads(result_text)
        except json.JSONDecodeError as je:
             print(f"JSON Parse Error: {je}")
             return {'actions': [], 'summary': "Error parsing AI response"}

        # Validate and Enhance
        if 'actions' in result and isinstance(result['actions'], list):
            enhanced_actions = []
            for action in result['actions']:
                # PYTHON-SIDE VALIDATION (Text Priority: NAME > EMAIL)
                # For text inputs ("Delete Virat"), the Name is the usage intent. The Email is likely AI hallucination.
                target_email = action.get('user_email', '').strip()
                target_name = action.get('user_name', '').strip()
                final_user = None
                
                # 1. Name Match (HIGHEST PRIORITY for Text)
                if target_name:
                     # Exact Name
                     matches = [u for u in user_list if u['user_name'].lower() == target_name.lower()]
                     if not matches:
                         # Fuzzy Name (e.g. "Virat" inside "Virat Kohli")
                         matches = [u for u in user_list if target_name.lower() in u['user_name'].lower() or u['user_name'].lower() in target_name.lower()]
                     
                     if matches:
                         # We found the user by Name! Trust this logic.
                         final_user = matches[0]
                
                # 2. Email Match (Fallback only if Name failed or wasn't provided)
                if not final_user and target_email:
                     matches = [u for u in user_list if u['email'].lower() == target_email.lower()]
                     if matches: 
                         # If we have a name ("Virat") but it didn't match anyone, 
                         # and the email ("salesforce@...") DOES match someone ("Salesforce Hub"),
                         # We have a CONFLICT. The user said "Virat", AI said "Salesforce".
                         # We should probably REJECT this to be safe, or log a warning.
                         # But if Name was empty, we trust Email.
                         if not target_name:
                             final_user = matches[0]
                         else:
                             # Name present but not found. Email present and valid.
                             # Likely AI hallucination. SAFEGUARD: Reject.
                             print(f"Skipping ambiguous action: Name '{target_name}' not found, but AI suggested Email '{target_email}' (User: {matches[0]['user_name']}). Mismatch suspected.")
                             final_user = None
                
                if final_user:
                    action['user_email'] = final_user['email']
                    action['user_name'] = final_user['user_name'] # Standardize
                    enhanced_actions.append(action)
            
            result['actions'] = enhanced_actions
        
        return result
    def _extract_json(self, text):
        """
        Robustly extract JSON object from text (handling markdown, comments, etc.)
        """
        text = text.strip()
        
        # Remove markdown code blocks if present
        if text.startswith('```'):
            # Find first newline to skip "```json"
            first_newline = text.find('\\n')
            if first_newline != -1:
                text = text[first_newline+1:]
            # Remove trailing "```"
            if text.endswith('```'):
                text = text[:-3]
        
        text = text.strip()
        
        # 1. Try to find the first outer-most brace pair
        start_idx = text.find('{')
        end_idx = text.rfind('}')
        
        if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
            return text[start_idx:end_idx+1]
            
        # 2. Try to find list brackets if object not found (for Excel actions)
        start_idx = text.find('[')
        end_idx = text.rfind(']')
        
        if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
            return text[start_idx:end_idx+1]
            
        return text

    def _enhance_action(self, action, user_list):
        """Enhance action with actual user data"""
        user_identifier = action.get('user_identifier', '').lower()
        
        # Try to find matching user
        # Try to find matching user with priority
        matched_user = None
        
        # 1. Exact Email Match
        for user in user_list:
            if user['email'].lower() == user_identifier:
                matched_user = user
                break
        
        # 2. Exact Name Match
        if not matched_user:
            for user in user_list:
                if user['user_name'].lower() == user_identifier:
                    matched_user = user
                    break
        
        # 3. Partial Match (Prioritize Longest Match)
        if not matched_user:
            best_match = None
            max_len = 0
            
            for user in user_list:
                u_name = user['user_name'].lower()
                u_email = user['email'].lower()
                
                # Check if identifier is contained in user data OR user data contained in identifier
                if (user_identifier in u_name or user_identifier in u_email or 
                    u_name in user_identifier): # name in identifier handles "Gaurav" vs "Gaurav Di"
                    
                    # Logic: "Gaurav" (6) in "Gaurav Di" (9) -> Match length 6
                    # "Gaurav Di" (9) in "Gaurav Di" (9) -> Match length 9 (BETTER)
                    
                    match_len = len(u_name)
                    if match_len > max_len:
                        max_len = match_len
                        best_match = user
            
            matched_user = best_match
        
        if matched_user:
            action['user_email'] = matched_user['email']
            action['user_name'] = matched_user['user_name']
            action['current_role'] = matched_user.get('roles', '')
        
        return action
    
    def parse_excel_actions(self, excel_data, user_list=None):
        """
        Parse actions from Excel data using AI with user context
        
        Args:
            excel_data: List of dicts from Excel rows
            user_list: List of valid users to match against
        
        Returns:
            Structured actions list
        """
        try:
            # Context string for AI
            user_context = ""
            context_instruction = ""
            
            if user_list:
                user_context = f"Valid Users List:\\n{json.dumps([{'name': u['user_name'], 'email': u['email'], 'role': u.get('roles', '')} for u in user_list], indent=2)}\\n"
                context_instruction = "2. **Data Integrity**: **TRUST THE EXCEL EMAIL** above all else."
            else:
                context_instruction = "2. **Data Integrity**: Use the exact Email and Name provided in the Excel Data columns."

            # Extract valid emails from Excel for strict enforcement
            valid_excel_emails = {row.get('Email', '').strip().lower() for row in excel_data if row.get('Email')}
            valid_emails_str = ", ".join(valid_excel_emails)

            prompt = f"""
You are a precise data processing assistant.
Your task is to convert the provided Excel Data into structured JSON actions.

{user_context}

Excel Data (Filtered - contains only relevant rows):
{json.dumps(excel_data, indent=2)}

INSTRUCTIONS:
1. **Row Isolation**: Process one row at a time. The 'Email' in the output MUST exactly match the 'Email' in that specific Excel row. NEVER copy an email from a previous row.
{context_instruction}
3. **Action Types**:
   - "delete" -> action_type: "delete"
   - "update to..." / "change role to..." -> action_type: "update_role"
   - "remove from..." -> action_type: "update_role" (unless "remove user" is explicit)
   - **CRITICAL**: If the 'Action' column is EMPTY, NULL, or just says "active", **IGNORE** this row. Do NOT generate an action.

4. **Column Mapping**:
   - `User Email`: COPY EXACTLY from the Excel row. Do not change a single character.
   - `Action`: Analyze the text to determine `action_type`.
   - `Details`: Use for `new_value` or context.

5. **STRICT CONSTRAINT**:
   - You are ONLY allowed to generate actions for these specific emails found in the source data: [{valid_emails_str}]
   - Do NOT hallucinate or substitute other emails from the "Valid Users List".

OUTPUT FORMAT (JSON Array):
[
  {{
    "action_type": "delete|update_role",
    "user_email": "EXACT_EMAIL_FROM_EXCEL_ROW",
    "user_name": "User name from Excel",
    "new_value": "Calculated final role string (for updates) or null",
    "description": "User [Name] deleted. OR Roles updated for [Name]."
  }}
]
IMPORTANT: Return ONLY the raw JSON string. Do not use Markdown code blocks.
"""
            
            # MODEL FALLBACK LOOP
            result_text = None
            last_error = None
            
            for model_name in self.models_to_try:
                try:
                    model = genai.GenerativeModel(model_name)
                    # Attempt generation
                    response = model.generate_content(prompt)
                    result_text = response.text.strip()
                    print(f"DEBUG: Success with model {model_name}")
                    print(f"DEBUG: RAW AI OUTPUT: {result_text}")
                    break # Success
                except exceptions.ResourceExhausted as e:
                    print(f"WARNING: Quota exceeded for {model_name}. Trying next...")
                    last_error = e
                    continue
                except Exception as e:
                    print(f"WARNING: Error with {model_name}: {e}. Trying next...")
                    last_error = e
                    continue
            
            if not result_text:
                print(f"Error parsing excel after trying all models. Last error: {last_error}")
                return []
            
            # Extract JSON
            json_match = re.search(r'\[.*\]', result_text, re.DOTALL)
            if json_match:
                result_text = json_match.group(0)
            
            actions = []
            try:
                actions = json.loads(result_text)
            except:
                # Fallback extraction
                pass

            # STRICT POST-PROCESSING: Filter out hallucinations
            # Only allow actions where the email matches a row in the input Excel file
            validated_actions = []
            for action in actions:
                email = action.get('user_email', '').strip().lower()
                if email in valid_excel_emails:
                    validated_actions.append(action)
                else:
                    print(f"WARNING: Blocked hallucinated email: {email} (Not in Excel source)")
            actions = validated_actions

            # Enhance actions with finding exact users from the list (Python-side double check)
            if user_list:
                enhanced_actions = []
                for action in actions:
                    # PYTHON-SIDE VALIDATION
                    
                    target_email = action.get('user_email', '').strip()
                    target_name = action.get('user_name', '').strip()
                    
                    final_user = None
                    
                    # 1. Try to find user by EMAIL first (Highest Priority)
                    if target_email:
                        matches = [u for u in user_list if u['email'].lower() == target_email.lower()]
                        if matches:
                            final_user = matches[0]
                    
                    # 2. If no valid email found, try to find by NAME
                    if not final_user and target_name:
                         # Try exact match first
                        matches = [u for u in user_list if u['user_name'].lower() == target_name.lower()]
                        if not matches:
                            # Try fuzzy/partial
                            matches = [u for u in user_list if target_name.lower() in u['user_name'].lower() or u['user_name'].lower() in target_name.lower()]
                        
                        if matches:
                            final_user = matches[0]
                    
                    # Apply Correction
                    if final_user:
                        action['user_email'] = final_user['email']
                        action['user_name'] = final_user['user_name']
                        action['current_role'] = final_user.get('roles', '')
                        enhanced_actions.append(action)
                    else:
                        # If we have an email but no user match, we keep it because it came from excel.
                        # Wait, what if it's a new user? Unlikely for "action" file.
                        if target_email:
                            enhanced_actions.append(action)
                return enhanced_actions
                    
            return actions
            
        except Exception as e:
            print(f"Error in parse_excel_actions: {e}")
            return []
    
    def test_connection(self):
        """Test Gemini API connection with Auto-Discovery"""
        last_error = None
        
        # Phase 1: Try Known Good Models
        for model_name in self.models_to_try:
            try:
                # print(f"Testing {model_name}...") 
                self.model_name = model_name
                model = genai.GenerativeModel(model_name)
                model.generate_content("Hello") # Simple check
                return True, f"Connection successful using model: {self.model_name}"
            except Exception as e:
                last_error = e
                continue
                
        # Phase 2: Auto-Discovery
        try:
            available_models = []
            valid_model_found = False
            
            for m in genai.list_models():
                if 'generateContent' in m.supported_generation_methods:
                    m_name = m.name.replace('models/', '')
                    available_models.append(m_name)
                    
                    if not valid_model_found:
                        try:
                            self.model_name = m_name
                            model = genai.GenerativeModel(m_name)
                            model.generate_content("Hello")
                            valid_model_found = True
                            return True, f"Auto-discovered working model: {m_name}"
                        except:
                            continue

            return False, f"Failed with standard models. Available: {available_models}. Last error: {str(last_error)}"
            
        except Exception as list_error:
            return False, f"Connection failed. Could not list models: {str(list_error)}. Original error: {str(last_error)}"
