# SpiroUI - Spirograph File Parser

Blazor WebAssembly application for parsing vintage Spirograph device files and converting them to JSON format.

## Features

- 🎨 Beautiful modern UI built with Tailwind CSS
- 📁 Drag & drop file upload support
- 🔄 Real-time file processing with progress indicators
- 📦 Batch processing of multiple files
- 💾 ZIP archive generation with parsed JSON files
- 🌐 Runs entirely in the browser - no backend required
- ⚡ Built with .NET 9 and Blazor WebAssembly

## Getting Started

### Prerequisites

- .NET 9 SDK
- Node.js (for Tailwind CSS)

### Installation

1. Clone the repository:
2. Install npm dependencies:
```bash
npm install
```

3. Build Tailwind CSS (for development):
```bash
npm run css:build
```

4. Run the application:
```bash
dotnet run
```

The application will be available at `https://localhost:5001` or `http://localhost:5000`.

### Building for Production

1. Build Tailwind CSS for production:
```bash
npm run css:prod
```

2. Build the Blazor application:
```bash
dotnet publish -c Release
```

The output will be in the `bin/Release/net9.0/publish/wwwroot` directory.

## Supported File Formats

The UI is configured to accept files with the following extensions:
- `.pnp`

You can modify the accepted file types in `Pages/Index.razor` by updating the `accept` attribute on the `InputFile` component.

## Architecture

- **Blazor WebAssembly**: Runs entirely in the browser using WebAssembly
- **Tailwind CSS**: Modern utility-first CSS framework for beautiful UI
- **System.IO.Compression**: Built-in .NET library for ZIP generation
- **JavaScript Interop**: For file downloads and enhanced drag-and-drop

## Development

### Tailwind CSS Development

The application uses Tailwind CSS with custom colors and animations. The configuration is in `tailwind.config.js`.

To watch for CSS changes during development:
```bash
npm run css:build
```

### Custom Colors

- `spiro-primary`: #6366f1 (Indigo)
- `spiro-secondary`: #8b5cf6 (Purple)
- `spiro-accent`: #ec4899 (Pink)

## Deployment

The application can be deployed as static files to any web server or CDN:

1. Build for production (see above)
2. Copy the contents of `wwwroot` to your web server
3. Configure your server to serve `index.html` for all routes (for client-side routing)

## License

MIT
