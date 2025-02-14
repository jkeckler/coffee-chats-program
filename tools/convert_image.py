from PIL import Image
import os

# Create assets directory if it doesn't exist
os.makedirs('assets', exist_ok=True)

# Convert image
img = Image.open('assets/NielsenIQ_logo_WhiteLetters_BlueBack.jpg')
img = img.resize((200, 200))  # Resize to match Excel template
img.save('assets/nielseniq_logo.png', 'PNG')