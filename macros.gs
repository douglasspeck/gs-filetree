const HOME = DriveApp.getFolderById('your-folder-id'); // Get the folder by ID (root of the shared drive)
const SHEET = SpreadsheetApp.getActiveSheet(); // Active sheet for output
const BUTTON = SpreadsheetApp.newTextStyle()
  .setBold(true) // Style for button text
  .setForegroundColor("#FF0000") // Red text color
  .build();

var depth = 0; // Tracks folder depth for visual indentation

function getAllFiles(folder) {
  // Recursive function to retrieve all files and subfolders in the given folder
  depth++; // Increment depth as we move into a folder

  let subfolders = folder.getFolders(); // Get subfolders
  let files = folder.getFiles(); // Get files
  let all_files = []; // Array to store file and folder data

  // Process each subfolder
  while (subfolders.hasNext()) {
    let next_folder = subfolders.next();
    all_files.push({
      name: next_folder.getName(),
      url: next_folder.getUrl(),
      depth: depth,
      has_next: subfolders.hasNext() || files.hasNext() // Check if more items exist
    });
    // Recursively gather all files in subfolders
    all_files.push(...getAllFiles(next_folder));
  }

  // Process each file in the current folder
  while (files.hasNext()) {
    let next_file = files.next();
    all_files.push({
      name: next_file.getName(),
      url: next_file.getUrl(),
      depth: depth,
      has_next: files.hasNext()
    });
  }

  depth--; // Decrease depth as we leave the folder
  return all_files; // Return the complete file/folder structure
}

function createFileTree(folder) {
  // Builds a tree-like structure for visualizing files and folders
  let files = getAllFiles(folder); // Retrieve all files and folders
  let tree = ""; // String to store the tree structure

  for (let i = 0; i < files.length; i++) {
    let file = files[i];
    file.prefix = ""; // Indentation or visual prefix

    // Add visual connectors based on folder depth
    if (file.depth > 1) {
      for (j = i - 1; j >= 0; j--) {
        if (files[j].depth == file.depth) {
          file.prefix = files[j].prefix;
          break;
        } else if (files[j].depth == file.depth - 1) {
          file.prefix = files[j].prefix;
          file.prefix += files[j].has_next ? "┃" : " "; // Add branching lines if more files exist
          break;
        }
      }
    }

    // Build the tree string with appropriate connectors (┣, ┗) and file/folder names
    tree += "\n" + file.prefix;
    tree += file.has_next ? "┣" : "┗";
    tree += file.name;
  }

  return [tree, files]; // Return the tree string and file list
}

function resizeSheet(sheet, max_rows, max_cols) {
  // Adjusts the sheet size to ensure there are enough rows and columns
  let cols = sheet.getMaxColumns();
  let rows = sheet.getMaxRows();
  if (cols < max_cols) {
    sheet.insertColumnsAfter(cols, max_cols - cols); // Inserts additional columns if necessary
  }
  if (rows < max_rows) {
    sheet.insertRowsAfter(rows, max_rows - rows); // Inserts additional rows if necessary
  }
}

function alignArray(a, n) {
  // Ensures that each row in the array has 'n' elements, filling with empty strings if needed
  a.forEach((r) => {
    for (let i=0; i<n; i++) {
      r[i] = r[i] ? r[i] : ""; // Fills missing elements with empty strings
    }
  });
  return a;
}

function createFileTreeSheet(folder) {
  // Creates the file tree structure in the spreadsheet
  let [tree, files] = createFileTree(folder); // Generate the file tree
  tree = tree.split("\n"); // Split the tree into lines
  resizeSheet(SHEET, files.length, 1); // Resize the sheet to fit the file structure

  for (let i = 1; i <= files.length; i++) {
    let row = files[i - 1]; // Get the current file or folder
    let prefix = tree[i].split(""); // Split the prefix for formatting
    tree[i] = prefix; // Passes the the splitted value to a variable out of this scope

    for (let j = 0; j < row.name.length; j++) { prefix.pop(); } // Adjust the prefix length to match the file name
    resizeSheet(SHEET, files.length, prefix.length + 1); // Resize sheet if more columns are needed

    // Set the prefix characters in the spreadsheet
    for (let j = 1; j <= prefix.length; j++) {
      SHEET.getRange(i, j).setValue(prefix[j - 1]);
    }
    
    // Create a rich text cell with a hyperlink to the file or folder
    let cell = SpreadsheetApp.newRichTextValue()
      .setText(row.name)
      .setLinkUrl(row.url)
      .setTextStyle(BUTTON)
      .build();
    SHEET.getRange(i, prefix.length + 1).setRichTextValue(cell); // Place the linked name in the sheet
  }

  // Merge the columns after the file names to create a cleaner view
  for (let i = 1; i <= files.length; i++) {
    let prefix = tree[i];
    SHEET.getRange(i, prefix.length + 1, 1, SHEET.getMaxColumns() - prefix.length + 1).merge();
  }

  SHEET.setColumnWidth(SHEET.getMaxColumns(), 300); // Set a fixed column width for the file/folder names
}

function cleanPage(sheet) {
  // Cleans the sheet by removing extra columns and rows, and clears content from cell A1
  let cols = sheet.getMaxColumns();
  let rows = sheet.getMaxRows();
  if (cols > 1) { sheet.deleteColumns(2, cols - 1); } // Remove excess columns
  if (rows > 1) { sheet.deleteRows(2, rows - 1); } // Remove excess rows
  sheet.getRange("A1").clearContent(); // Clear content from cell A1
}

function scan() {
  // Entry point to clear the sheet and create the file tree
  cleanPage(SHEET); // Clean the sheet before generating the tree
  createFileTreeSheet(HOME); // Generate the file tree from the root folder
}

function onLoad() {
  // Adds a custom menu to the spreadsheet for generating the file tree
  const UI = SpreadsheetApp.getUi(); // UI instance for dialog interactions
  let menu = UI.createMenu('Tree');
  menu.addItem('New File Tree', 'scan'); // Adds a menu item to call the 'scan' function
  menu.addToUi(); // Adds the menu to the spreadsheet UI
}