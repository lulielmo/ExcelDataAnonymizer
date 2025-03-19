import pandas as pd
import random
import string
from pathlib import Path
import json
import re

class ExcelAnonymizer:
    def __init__(self):
        """Initierar anonymizer."""
        self.name_mapping = {}  # Fullständigt namn -> (förnamnsalias, efternamnsalias)
        self.email_mapping = {}  # Original e-post -> anonymiserad e-post
        self.first_name_mapping = {}  # Förnamn -> förnamnsalias
        self.last_name_mapping = {}  # Efternamn -> efternamnsalias
        self.username_mapping = {}  # Användarnamn -> anonymiserat användarnamn
        
    def generate_alias(self):
        """Genererar ett slumpmässigt alias för ett namn."""
        return ''.join(random.choices(string.ascii_letters, k=8))
    
    def extract_name_from_email(self, email):
        """Extraherar namn från en e-postadress."""
        if not isinstance(email, str) or '@' not in email:
            return None
        username = email.split('@')[0]
        # Hantera både punkt och bindestreck som separatorer
        name_parts = re.split('[.-]', username)
        if len(name_parts) >= 2:
            first_name = name_parts[0]
            last_name = name_parts[1]
            return f"{first_name} {last_name}"
        return None
    
    def anonymize_full_name(self, full_name):
        """Anonymiserar ett fullständigt namn och behåller mappningen."""
        if not isinstance(full_name, str):
            return full_name
            
        # Om namnet redan är anonymiserat, returnera det existerande alias
        if full_name in self.name_mapping:
            first_alias, last_alias = self.name_mapping[full_name]
            return f"{first_alias} {last_alias}"
            
        # Dela upp namnet i förnamn och efternamn
        name_parts = full_name.split()
        if len(name_parts) < 2:
            return full_name
            
        first_name = name_parts[0]
        last_name = ' '.join(name_parts[1:])
        
        # Skapa nya alias för för- och efternamn
        first_alias = self.generate_alias()
        last_alias = self.generate_alias()
        
        # Spara mappningar
        self.name_mapping[full_name] = (first_alias, last_alias)
        self.first_name_mapping[first_name] = first_alias
        self.last_name_mapping[last_name] = last_alias
        
        return f"{first_alias} {last_alias}"
    
    def anonymize_username(self, username, full_name=None):
        """Anonymiserar ett användarnamn."""
        if not isinstance(username, str):
            return username
            
        if username in self.username_mapping:
            return self.username_mapping[username]
            
        if full_name and full_name in self.name_mapping:
            first_alias, last_alias = self.name_mapping[full_name]
            new_username = f"{first_alias}.{last_alias}"
        else:
            # Om vi inte har ett namn, skapa nya alias
            first_alias = self.generate_alias()
            last_alias = self.generate_alias()
            new_username = f"{first_alias}.{last_alias}"
            
            # Om vi har ett användarnamn med punkt, spara mappningen för framtida användning
            if '.' in username:
                name_parts = username.split('.')
                if len(name_parts) >= 2:
                    constructed_name = f"{name_parts[0]} {name_parts[1]}"
                    self.name_mapping[constructed_name] = (first_alias, last_alias)
        
        self.username_mapping[username] = new_username
        return new_username
    
    def anonymize_email(self, email):
        """Anonymiserar en e-postadress baserat på namnmappningen."""
        if not isinstance(email, str) or '@' not in email:
            return email
            
        try:
            # Dela upp e-postadressen vid första @-tecknet
            parts = email.split('@', 1)
            if len(parts) != 2:
                return email
                
            username = parts[0]
            domain = parts[1]
            
            # Försök hitta ett existerande namn från e-postadressen
            extracted_name = self.extract_name_from_email(email)
            if extracted_name and extracted_name in self.name_mapping:
                first_alias, last_alias = self.name_mapping[extracted_name]
            elif username in self.username_mapping:
                # Om vi redan har anonymiserat användarnamnet, använd samma alias
                username_alias = self.username_mapping[username]
                first_alias, last_alias = username_alias.split('.')
            else:
                # Skapa nya alias
                first_alias = self.generate_alias()
                last_alias = self.generate_alias()
                if extracted_name:
                    self.name_mapping[extracted_name] = (first_alias, last_alias)
            
            new_email = f"{first_alias}.{last_alias}@{domain}"
            self.email_mapping[email] = new_email
            return new_email
        except Exception as e:
            print(f"Varning: Kunde inte anonymisera e-postadressen: {email}")
            return email
    
    def anonymize_text(self, text):
        """Anonymiserar text genom att ersätta personnamn med alias."""
        if not isinstance(text, str):
            return text
            
        result = text
        
        # Ersätt e-postadresser först
        for original_email, anonymized_email in self.email_mapping.items():
            result = result.replace(original_email, anonymized_email)
        
        # Ersätt användarnamn
        for original_username, anonymized_username in self.username_mapping.items():
            result = result.replace(original_username, anonymized_username)
        
        # Ersätt fullständiga namn
        for original_name, (first_alias, last_alias) in self.name_mapping.items():
            result = result.replace(original_name, f"{first_alias} {last_alias}")
            # Ersätt även punkt-separerad version
            dot_version = original_name.replace(' ', '.')
            result = result.replace(dot_version, f"{first_alias}.{last_alias}")
            
        return result
    
    def anonymize_excel(self, input_file, output_file):
        """
        Anonymiserar en Excel-fil genom att ersätta personnamn, e-postadresser och användarnamn.
        """
        # Läs Excel-filen utan att skippa rader
        df = pd.read_excel(input_file)
        
        # Definiera standardvärden som inte ska anonymiseras
        standard_values = {
            'Access License Type',
            'Teammedlemmar',
            'Inget',
            'System user',
            'Mobility user',
            'Security Role',
            'Medius Adapter',
            'Alias',
            'Operations',
            'Aktivitet',
            'Användarnamn',
            'Nätverksdomän'
        }
        
        # Hitta kolumner som innehåller "Alias" och "Användarnamn"
        alias_column = None
        username_column = None
        
        # Gå igenom alla kolumner för att hitta de som innehåller Alias och Användarnamn
        for col in df.columns:
            col_values = df[col].astype(str).str.strip()
            if "Alias" in col_values.values:
                alias_column = col
            if "Användarnamn" in col_values.values:
                username_column = col
        
        if alias_column is None or username_column is None:
            print("Varning: Kunde inte hitta både Alias- och Användarnamn-kolumner")
            return 0
            
        print(f"Hittade följande kolumner att anonymisera:")
        print(f"Alias-kolumn: {alias_column}")
        print(f"Användarnamn-kolumn: {username_column}")
        
        # Skapa en kopia av dataframe för att undvika varningar
        df_copy = df.copy()
        
        # Anonymisera värden i Alias-kolumnen
        for idx in range(len(df)):
            value = df.at[idx, alias_column]
            if not isinstance(value, str) or not value.strip() or value.strip() in standard_values:
                continue
                
            if '@' in value:
                df_copy.at[idx, alias_column] = self.anonymize_email(value)
            else:
                df_copy.at[idx, alias_column] = self.anonymize_username(value)
        
        # Anonymisera värden i Användarnamn-kolumnen
        for idx in range(len(df)):
            value = df.at[idx, username_column]
            if not isinstance(value, str) or not value.strip() or value.strip() in standard_values:
                continue
            df_copy.at[idx, username_column] = self.anonymize_username(value)
        
        # Spara den anonymiserade datan
        df_copy.to_excel(output_file, index=False)
        
        # Konvertera mappningar till önskat format för JSON
        formatted_name_mapping = {
            name: f"{first} {last}" 
            for name, (first, last) in self.name_mapping.items()
        }
        
        formatted_email_mapping = {
            email: new_email
            for email, new_email in self.email_mapping.items()
        }
        
        formatted_username_mapping = {
            username: new_username
            for username, new_username in self.username_mapping.items()
        }
        
        # Spara mappningarna
        mapping_file = Path(output_file).with_suffix('.mapping.json')
        mapping_data = {
            'name_mapping': formatted_name_mapping,
            'email_mapping': formatted_email_mapping,
            'username_mapping': formatted_username_mapping
        }
        with open(mapping_file, 'w', encoding='utf-8') as f:
            json.dump(mapping_data, f, ensure_ascii=False, indent=2)
        
        total_anonymized = len(self.name_mapping) + len(self.email_mapping) + len(self.username_mapping)
        return total_anonymized

def main():
    """Huvudfunktion för att köra anonymiseringen."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Anonymiserar personnamn i Excel-filer')
    parser.add_argument('input_file', help='Sökväg till indatafilen')
    parser.add_argument('output_file', help='Sökväg till utdatafilen')
    
    args = parser.parse_args()
    
    anonymizer = ExcelAnonymizer()
    num_anonymized = anonymizer.anonymize_excel(
        args.input_file,
        args.output_file
    )
    
    print(f"Anonymisering klar! {num_anonymized} namn har ersatts.")
    print(f"Resultatet har sparats i: {args.output_file}")
    print(f"Mappningen har sparats i: {Path(args.output_file).with_suffix('.mapping.json')}")

if __name__ == "__main__":
    main() 