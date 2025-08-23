# Entry point wrapper for deployment platforms expecting a specific filename
# Delegates to main() in app.py
from app import main

if __name__ == "__main__":
    main()
