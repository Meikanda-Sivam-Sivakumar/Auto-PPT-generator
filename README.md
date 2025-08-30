#  AI PowerPoint Generator

**Created by Meikanda Sivam**

A professional AI-powered PowerPoint presentation generator that converts text content into beautifully formatted presentations using OpenAI, Anthropic, or Groq AI models.

##  Features

- **Multiple AI Providers**: Support for OpenAI GPT, Anthropic Claude, and Groq Mixtral
- **Custom Templates**: Upload your own PowerPoint templates for brand consistency
- **Speaker Notes**: Automatically generate detailed presenter notes
- **Professional UI**: Clean, corporate-grade interface
- **Secure**: API keys are processed securely and never stored
- **Fast Processing**: Optimized for quick generation and download

##  Deployment on Render

### Prerequisites
- A Render account (https://render.com)
- Git repository with this code

### Deployment Steps

1. **Fork/Clone this repository** to your GitHub account

2. **Create a new Web Service** on Render:
   - Connect your GitHub repository
   - Set the build command: `pip install -r requirements.txt`
   - Set the start command: `gunicorn app:app`
   - Choose Python 3.11 runtime

3. **Environment Variables** (Optional):
   - `FLASK_ENV=production`
   - `PORT` (automatically set by Render)

4. **Deploy**: Render will automatically build and deploy your application

## 🛠 Technology Stack

- **Backend**: Flask (Python)
- **Frontend**: Bootstrap 5, HTML5, CSS3
- **AI Integration**: OpenAI, Anthropic, Groq APIs
- **Document Generation**: python-pptx
- **Deployment**: Render-ready configuration

##  Quick Start

### Prerequisites
- Python 3.8 or higher
- A web browser

### Installation & Setup

1. **Clone or Download** this repository to your computer

2. **Get an API Key** from one of these providers:
   - **OpenAI**: Visit [platform.openai.com/api-keys](https://platform.openai.com/api-keys)
   - **Anthropic**: Visit [console.anthropic.com/account/keys](https://console.anthropic.com/account/keys)
   - **Groq**: Visit [console.groq.com/keys](https://console.groq.com/keys)

3. **Run the Application**:

   **On Windows:**
   ```bash
   double-click start.bat
   ```

   **On Mac/Linux:**
   ```bash
   chmod +x start.sh
   ./start.sh
   ```

   **Manual Setup:**
   ```bash
   cd backend
   python -m venv venv
   
   # On Windows:
   venv\Scripts\activate
   
   # On Mac/Linux:
   source venv/bin/activate
   
   pip install -r requirements.txt
   python app.py
   ```

4. **Open the Frontend**:
   - Open `frontend/index.html` in your web browser
   - Or visit the direct file path in your browser

##  How to Use

1. **Enter Your Content**: Paste the text you want to convert into slides
2. **Choose AI Provider**: Select OpenAI, Anthropic, or Groq
3. **Enter API Key**: Input your API key (it's not stored)
4. **Optional Guidance**: Add specific instructions like "make it formal" or "keep it simple"
5. **Generate**: Click the generate button and wait
6. **Download**: Your PowerPoint file will automatically download

## 📖 Usage Examples

### Business Report
```
Input: "Our Q3 sales increased by 25% compared to Q2. The main drivers were improved marketing campaigns and new product launches..."

Output: Professional slides with:
- Title slide
- Executive Summary
- Key Metrics
- Growth Drivers
- Conclusion
```

### Educational Content
```
Input: "Photosynthesis is the process by which plants convert sunlight into energy. The process occurs in chloroplasts..."

Guidance: "Make it suitable for high school students"

Output: Educational slides with:
- Introduction to Photosynthesis
- The Process Explained
- Key Components
- Importance in Nature
```

## 🔧 API Endpoints

The backend provides these endpoints:

- `GET /health` - Health check
- `GET /providers` - List supported AI providers
- `POST /generate` - Generate PowerPoint presentation

### Generate Endpoint
```json
{
  "text": "Your content here",
  "provider": "openai|anthropic|groq",
  "api_key": "your-api-key",
  "guidance": "optional styling guidance"
}
```

## 🔒 Security & Privacy

- **API Keys**: Never stored or logged, only used for the generation request
- **Content**: Your text content is only sent to your chosen AI provider
- **Files**: Generated PowerPoint files are temporary and not stored on the server
- **Local**: The application runs locally on your machine

## 🛠️ Technical Details

### Backend Stack
- **Flask**: Web framework
- **python-pptx**: PowerPoint generation
- **OpenAI/Anthropic/Groq SDKs**: AI integration
- **Flask-CORS**: Cross-origin requests

### Frontend Stack
- **HTML/CSS/JavaScript**: Simple web interface
- **Bootstrap 5**: UI framework
- **Font Awesome**: Icons

### Project Structure
```
PPT-generator/

│── app.py              # Main Flask application
│── requirements.txt    # Python dependencies
│── index.html         # Web interface           # Mac/Linux startup script
└── README.md              # This file
```

## 🔧 Troubleshooting

### Common Issues

**"Python is not installed"**
- Install Python 3.8+ from [python.org](https://python.org)

**"API key is invalid"**
- Check that your API key is correct and has sufficient credits
- Ensure you're using the right provider (OpenAI keys start with "sk-", Anthropic with "sk-ant-", Groq with "gsk_")

**"CORS errors"**
- Make sure the backend server is running on port 5000
- Try refreshing the frontend page

**"Generation failed"**
- Check your internet connection
- Verify your API key has sufficient credits
- Try with shorter text content

**"Download doesn't start"**
- Check your browser's download settings
- Try a different browser
- Ensure popup blockers aren't interfering

### Getting Help

1. Check the browser console for error messages (F12 → Console)
2. Check the backend terminal for error logs
3. Try with different AI providers
4. Ensure your content isn't too long (< 10,000 characters recommended)

## 🎨 Customization

### Adding New AI Providers
1. Install the provider's SDK in `requirements.txt`
2. Add the provider logic in the `LLMOrchestrator` class
3. Update the frontend provider cards

### Modifying Slide Templates
Edit the `PPTGenerator` class in `app.py` to customize:
- Slide layouts
- Font sizes and colors
- Bullet point styles
- Title formatting

### Styling the Interface
Modify `frontend/index.html` to change:
- Colors and themes
- Layout and spacing
- Button styles
- Form elements

## 📝 License

This project is provided as-is for educational and personal use. Feel free to modify and distribute as needed.

## 🤝 Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool!

---

**Enjoy creating amazing presentations with AI! 🎉**
