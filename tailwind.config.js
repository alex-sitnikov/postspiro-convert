/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./**/*.{razor,html,cshtml}",
    "./wwwroot/index.html"
  ],
  theme: {
    extend: {
      colors: {
        'spiro-primary': '#6366f1',
        'spiro-secondary': '#8b5cf6',
        'spiro-accent': '#ec4899',
      },
      animation: {
        'pulse-slow': 'pulse 3s cubic-bezier(0.4, 0, 0.6, 1) infinite',
        'float': 'float 3s ease-in-out infinite',
        'spin-slow': 'spin 3s linear infinite',
      },
      keyframes: {
        float: {
          '0%, 100%': { transform: 'translateY(0)' },
          '50%': { transform: 'translateY(-10px)' },
        }
      },
      scale: {
        '102': '1.02',
      }
    },
  },
  plugins: [],
}