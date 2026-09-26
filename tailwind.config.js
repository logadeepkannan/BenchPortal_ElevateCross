/** @type {import('tailwindcss').Config} */
export default {
  darkMode: "class",
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      keyframes: {
        "clippy-float": {
          "0%, 100%": { transform: "translateY(0px) rotate(-1deg)" },
          "50%": { transform: "translateY(-8px) rotate(1deg)" },
        },
        "clippy-blink": {
          "0%, 90%, 100%": { transform: "scaleY(1)" },
          "95%": { transform: "scaleY(0.1)" },
        },
        "clippy-eyebrow": {
          "0%, 100%": { transform: "translateY(0px) rotate(0deg)" },
          "50%": { transform: "translateY(-1.5px) rotate(-3deg)" },
        },
        "clippy-think": {
          "0%, 100%": { transform: "rotate(-6deg)" },
          "25%": { transform: "rotate(6deg)" },
          "50%": { transform: "rotate(-8deg)" },
          "75%": { transform: "rotate(8deg)" },
        },
        "clippy-bounce": {
          "0%, 100%": { transform: "translateY(0) scale(1)" },
          "30%": { transform: "translateY(-18px) scale(1.05)" },
          "50%": { transform: "translateY(0) scale(0.96)" },
          "70%": { transform: "translateY(-8px) scale(1.02)" },
        },
        "pop-in": {
          "0%": { opacity: "0", transform: "scale(0.9) translateY(6px)" },
          "100%": { opacity: "1", transform: "scale(1) translateY(0)" },
        },
        "fade-in": {
          "0%": { opacity: "0" },
          "100%": { opacity: "1" },
        },
      },
      animation: {
        "clippy-float": "clippy-float 3.5s ease-in-out infinite",
        "clippy-blink": "clippy-blink 4.5s ease-in-out infinite",
        "clippy-eyebrow": "clippy-eyebrow 2.2s ease-in-out infinite",
        "clippy-think": "clippy-think 0.35s ease-in-out infinite",
        "clippy-bounce": "clippy-bounce 0.7s ease-in-out 2",
        "pop-in": "pop-in 0.18s ease-out",
        "fade-in": "fade-in 0.25s ease-out",
      },
    },
  },
  plugins: [],
};
