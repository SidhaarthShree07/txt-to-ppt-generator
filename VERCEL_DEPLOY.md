# Quick Vercel Deployment Guide

This guide will help you deploy the Text-to-PowerPoint Generator to Vercel in just a few minutes.

## Method 1: One-Click Deploy (Recommended)

1. **Click the Deploy Button**:
   [![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new/clone?repository-url=https://github.com/SidhaarthShree07/txt-to-ppt-generator)

2. **Follow the prompts**:
   - Connect your GitHub account
   - Choose a repository name
   - Click "Deploy"

3. **Wait for deployment** (usually takes 2-3 minutes)

4. **Your app is live!** 🎉

## Method 2: Vercel CLI

1. **Install Vercel CLI**:
   ```bash
   npm install -g vercel
   ```

2. **Clone the repository**:
   ```bash
   git clone https://github.com/SidhaarthShree07/txt-to-ppt-generator.git
   cd text-to-powerpoint
   ```

3. **Login to Vercel**:
   ```bash
   vercel login
   ```

4. **Deploy**:
   ```bash
   vercel --prod
   ```

5. **Follow the prompts** and your app will be deployed!

## Method 3: GitHub Integration

1. **Fork this repository** to your GitHub account

2. **Go to [vercel.com](https://vercel.com)** and sign in

3. **Click "New Project"**

4. **Import your forked repository**

5. **Click "Deploy"** - Vercel will automatically detect the Flask app

## What's Included for Vercel

The repository already includes all necessary Vercel configuration files:

- ✅ `vercel.json` - Vercel configuration
- ✅ `runtime.txt` - Python version specification  
- ✅ `requirements.txt` - Python dependencies
- ✅ `.vercelignore` - Files to exclude from deployment

## Important Notes

### Serverless Function Limits
- **Hobby Plan**: 10-second timeout limit
- **Pro Plan**: 30-second timeout limit
- Large presentations might hit timeout limits on Hobby plan

### Memory Limits
- **Hobby Plan**: 1024 MB memory
- **Pro Plan**: 3008 MB memory

### File Processing
- Temporary files are handled automatically
- No persistent storage (files are cleaned up after processing)

## Testing Your Deployment

1. **Visit your deployed URL**
2. **Get a Gemini API key**: [Google AI Studio](https://makersuite.google.com/app/apikey)
3. **Test with sample content**:
   ```
   # Sample Text
   Our company has seen tremendous growth this year. We've expanded our team by 50% and increased revenue by 300%. Our key achievements include launching three new products, entering five new markets, and establishing partnerships with major industry players.
   
   Looking ahead, we plan to continue this momentum with strategic investments in technology and talent acquisition.
   ```

## Troubleshooting

### Common Issues

**Deployment Failed**:
- Check that all files are properly committed to your repository
- Ensure `requirements.txt` includes all dependencies

**Timeout Errors**:
- Large PowerPoint templates may cause timeouts on Hobby plan
- Consider using smaller templates or upgrading to Pro plan

**Import Errors**:
- Vercel automatically handles Python path configuration
- All dependencies should be listed in `requirements.txt`

**API Errors**:
- Ensure you're using a valid Gemini API key
- Check API key format and permissions

### Getting Help

- Check the [Vercel Documentation](https://vercel.com/docs)
- Open an issue in the GitHub repository
- Check Vercel deployment logs for specific errors

## Custom Domain (Optional)

1. **In your Vercel dashboard**, go to your project
2. **Click "Domains"** tab
3. **Add your custom domain**
4. **Follow DNS setup instructions**

## Environment Variables

This app doesn't require any environment variables since users provide their own API keys. However, if you want to add any:

1. **In your Vercel dashboard**, go to project settings
2. **Click "Environment Variables"**
3. **Add any needed variables**

## Monitoring

- **View logs**: Vercel Dashboard → Your Project → Functions tab
- **Monitor usage**: Vercel Dashboard → Your Project → Analytics
- **Health check**: Visit `your-domain.com/api/health`

## Next Steps

Once deployed:

1. **Share your app** with others
2. **Customize the branding** in the HTML templates
3. **Add analytics** (Google Analytics, Vercel Analytics)
4. **Consider upgrading** to Pro plan for better performance

---

🚀 **That's it!** Your Text-to-PowerPoint Generator is now live on Vercel and ready to use!
