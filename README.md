# Transfer Azure DevOps Tests

This repo serves as a tool to batch transfer Test Cases from one DevOps enviornment to another when the environments are not shared.
DevOps allows for this same function between shared environments, but otherwise, it is a manual click-instensive operation.

This solution uses a macro (Recorded and run by TinyTask) to download Test Cases in bunches.
Then, a small C# CLI takes the Test Case download directory and concatenates and reformats the `.csv` files into a single, DevOps import friendly file.

### Set up browser 🌐
- Login to DevOps and navigate to the desired Test Suite
- Go to browser settings and relocate the downloads directory to a fresh directory for this operation (Change it back later)
- Record the macro
- Test the macro
- Run the macro until all desired Test Cases are downloaded

### Run the CLI 💾
- Run the program and follow the prompts
- Verify `output.csv`

### Import 🔼
- Use the import function in the Test Suite browser and select `output.csv`
- This operation will take longer with the more test cases included in the file

## Known Bugs 🐞
- DevOps seems to assign at least one Test Case to the user that uploads `output.csv` even though the field is empty in the file
- Some test cases get duplicated by DevOps
- Unknown syntax parsing error by DevOps
