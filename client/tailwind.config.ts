import type { Config } from 'tailwindcss';

export default {
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        ink: '#0f1720',
        sand: '#efe6d2',
        fairway: '#1e5f42',
        pine: '#163d2c',
        mist: '#f5f3ee',
        gold: '#d0a95f',
        clay: '#b65d43',
      },
      boxShadow: {
        panel: '0 24px 80px rgba(15, 23, 32, 0.12)',
      },
      backgroundImage: {
        hero: 'radial-gradient(circle at top left, rgba(208, 169, 95, 0.32), transparent 32%), linear-gradient(135deg, #163d2c 0%, #1e5f42 45%, #efe6d2 160%)',
      },
      fontFamily: {
        sans: ['Manrope', 'sans-serif'],
        display: ['"DM Serif Display"', 'serif'],
      },
    },
  },
  plugins: [],
} satisfies Config;
