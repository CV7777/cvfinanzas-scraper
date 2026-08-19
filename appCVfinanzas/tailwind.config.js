module.exports = {
  purge: ['./src/views/**/*.ejs', './public/**/*.html', './public/js/**/*.js', './src/styles/**/*.css'],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        primary: '#4a7c9e',
        secondary: '#e05a7a',
      },
    },
  },
  variants: {
    extend: {},
  },
  plugins: [],
};
