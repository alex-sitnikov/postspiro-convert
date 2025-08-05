#!/bin/bash

echo "🚀 Deploying SpiroUI to GitHub Pages..."

# Build for GitHub Pages
echo "📦 Building for production..."
npm run css:prod
dotnet publish -c Release -o gh-pages-dist

# Update base href for GitHub Pages
echo "🔧 Updating base href for GitHub Pages..."
sed -i 's|<base href="/" />|<base href="/SpiroUI/" />|g' gh-pages-dist/wwwroot/index.html

echo "✅ Build complete!"
echo "📁 Output directory: gh-pages-dist/wwwroot"
echo ""
echo "📝 Next steps:"
echo "1. Push your code to GitHub"
echo "2. Go to repository Settings → Pages"
echo "3. Set source to 'GitHub Actions'"
echo "4. The workflow will automatically deploy on push to main"
echo ""
echo "🌐 Your app will be available at:"
echo "   https://[your-username].github.io/SpiroUI/"