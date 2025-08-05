# GitHub Pages Deployment Guide

## 🚀 Automatic Deployment (Recommended)

### 1. Push to GitHub
```bash
git add .
git commit -m "Add GitHub Pages deployment"
git push origin main
```

### 2. Enable GitHub Pages
1. Go to your repository on GitHub
2. Click **Settings** → **Pages**
3. Under **Source**, select **GitHub Actions**
4. The workflow will automatically deploy on every push to main

### 3. Access Your App
Your app will be available at:
```
https://[your-username].github.io/SpiroUI/
```

## 🛠️ Manual Deployment

If you prefer to deploy manually:

```bash
./deploy-github-pages.sh
```

Then upload the contents of `gh-pages-dist/wwwroot` to your GitHub Pages.

## 📁 Deployment Files Created

- **`.github/workflows/deploy.yml`** - GitHub Actions workflow
- **`wwwroot/.nojekyll`** - Prevents Jekyll processing
- **`wwwroot/404.html`** - Handles client-side routing
- **`deploy-github-pages.sh`** - Manual deployment script

## 🔧 How It Works

1. **GitHub Actions Workflow** runs on every push to main
2. **Installs .NET 9** and Node.js
3. **Builds Tailwind CSS** with production optimizations
4. **Publishes Blazor WebAssembly** app
5. **Deploys to GitHub Pages** automatically

## ⚡ Features

- ✅ Automatic deployment on push
- ✅ Tailwind CSS build optimization
- ✅ Client-side routing support
- ✅ Proper base path handling
- ✅ No Jekyll interference
- ✅ Free HTTPS hosting

## 🐛 Troubleshooting

### Deployment Fails
- Check the **Actions** tab in your GitHub repository
- Ensure you have Pages enabled in repository settings
- Verify the workflow has proper permissions

### App Doesn't Load
- Check browser console for 404 errors
- Verify base href is correct in deployed index.html
- Ensure all assets are loading with the correct path

### Routing Issues
- The 404.html file handles client-side routing
- Make sure .nojekyll file exists to prevent Jekyll processing

## 📝 Custom Domain (Optional)

To use a custom domain:

1. Add a `CNAME` file to `wwwroot/` with your domain
2. Configure DNS to point to `[username].github.io`
3. Update base href in index.html if needed

## 🔄 Updates

To update your deployed app:
1. Make changes to your code
2. Push to the main branch
3. GitHub Actions will automatically redeploy