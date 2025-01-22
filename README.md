# Giftcard Automation Script

This repository contains a script to automate the creation of personalized gift cards using Adobe Photoshop. The script processes a CSV file with gift card details, generates designs, and exports them as PNG and PSD files.

## Features

- Reads a CSV file containing gift card details (e.g., recipient, message, sender, and order ID).
- Automatically places text and images into a pre-defined gift card template.
- Exports the final designs as both PNG and PSD files.
- Organizes outputs into time-stamped folders for easy management.

## How to Use

1. **Prepare Your Files**:
   - Place the `Artkive_GC.png` (logo) and `Giftcards.png` (background) in the same folder as the script.
   - Ensure your CSV file is formatted with columns for `To`, `Message`, `From`, and `Order ID`.

2. **Run the Script**:
   - Double-click `000 RUN ME with powershell to start GIFTCARD automation.ps1`.
   - The script will process the CSV file and create the gift cards.

3. **Check the Output**:
   - Generated gift cards will be saved in the `Output` folder under a time-stamped subfolder.

## Requirements

- Adobe Photoshop.
- A properly formatted CSV file with gift card data.

## Notes

- Test the script with a small CSV file before processing large batches.
- Ensure Adobe Photoshop has permissions to execute scripts.

## Contributing

Feel free to fork this repository and suggest improvements or report any issues.

## License

This project is open-source and available under the [MIT License](LICENSE).
