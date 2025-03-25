// i want to create a whatsapp bot from scratch using javascript. I am using vscode as editor. Now here is what i wanna do: I have a excel sheet (a_data.xlsx) contain 4 columns: Name, Place, Age, Salary. now this excel sheet should be connected to chat bot. when i text hi to bot it should reply with " Hi, please enter your name". and when user enter his name it should search in name column and if it finds the name it should return all the details against it like age salary & place of that person. Then it should ask if user wants to change anything (with option of yes & no) & if user say yes then it should ask which value it wants to change (give options of age, place, salary). after choosing option user should enter new value. bot should update excel sheet with new value. then it should ask again if user wants to change anything till user say no. if user say no then it should end the chat with simple "thankyou, if you need further assistance say hi"

// https://jindalstainless0-my.sharepoint.com/:x:/g/personal/lucky1_jindalstainless_com/EQe6CWtVC1RFv9NClNFSsm8Btv3U7q3o45cpQorp5-QYsg?e=oJDlDn

// const { Client, LocalAuth } = require("whatsapp-web.js");
// const qrcode = require("qrcode-terminal");
// const xlsx = require("xlsx");
// const axios = require("axios");
// const fs = require("fs");


// // Shared OneDrive link (with edit permissions)
// const oneDriveSharedLink = "https://jindalstainless0-my.sharepoint.com/personal/lucky1_jindalstainless_com/_layouts/15/download.aspx?share=EQe6CWtVC1RFv9NClNFSsm8Btv3U7q3o45cpQorp5-QYsg"; 
// // Replace with your shared link

// // Path to the local Excel file
// const filePath = "./a_data.xlsx";

// // Global variables
// let workbook;
// let sheetName;
// let worksheet;
// let headerRow;
// let data = []; // Initialize data as an empty array

// // Function to download the file from OneDrive
// async function downloadFileFromOneDrive() {
//   try {
//     const response = await axios.get(oneDriveSharedLink, { responseType: "arraybuffer" });
//     fs.writeFileSync(filePath, response.data); // Save the file locally
//     console.log("File downloaded from OneDrive.");
//   } catch (error) {
//     console.error("Error downloading file:", error.message);
//   }
// }

// // Function to upload the file to OneDrive
// async function uploadFileToOneDrive() {
//   try {
//     const fileContent = fs.readFileSync(filePath); // Read the local file
//     await axios.put(oneDriveSharedLink, fileContent, {
//       headers: {
//         "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
//       },
//     });
//     console.log("File uploaded to OneDrive.");
//   } catch (error) {
//     console.error("Error uploading file:", error.message);
//   }
// }

// // Function to load the Excel file
// async function loadExcelFile() {
//   try {
//     // Download the file from OneDrive
//     await downloadFileFromOneDrive();

//     // Check if the file exists locally
//     if (!fs.existsSync(filePath)) {
//       console.error("File not found locally. Creating an empty Excel file.");
//       workbook = xlsx.utils.book_new(); // Create a new workbook
//       xlsx.utils.book_append_sheet(workbook, xlsx.utils.aoa_to_sheet([[]]), "Sheet1"); // Add an empty sheet
//       xlsx.writeFile(workbook, filePath); // Save the empty workbook
//     }

//     // Load the Excel file
//     workbook = xlsx.readFile(filePath);
//     sheetName = workbook.SheetNames[0];
//     worksheet = workbook.Sheets[sheetName];
//     headerRow = xlsx.utils.sheet_to_json(worksheet, { header: 1 })[0]; // Get the header row

//     data = xlsx.utils.sheet_to_json(worksheet);

//     // Normalize data fields, ensuring all header fields are present
//     data = data.map((entry) => {
//       let normalizedEntry = {};
//       headerRow.forEach((header) => {
//         const trimmedHeader = header?.toString().trim(); // Handle potential undefined header
//         normalizedEntry[trimmedHeader] = entry[trimmedHeader] || "N/A"; // Use value or "N/A"
//       });
//       return normalizedEntry;
//     });

//     console.log("📦 Loaded data from local file:", JSON.stringify(data, null, 2));
//   } catch (error) {
//     console.error("Error loading Excel file:", error.message);
//   }
// }

// // Function to update Excel file
// async function updateExcelFile() {
//   try {
//     const updatedWorksheet = xlsx.utils.json_to_sheet(data);
//     workbook.Sheets[sheetName] = updatedWorksheet;
//     xlsx.writeFile(workbook, filePath);
//     console.log("Excel file updated locally.");
//     await uploadFileToOneDrive(); // Upload the updated file to OneDrive
//   } catch (error) {
//     console.error("Error updating Excel file:", error);
//   }
// }

// // Initialize WhatsApp client
// const client = new Client({
//   authStrategy: new LocalAuth(),
//   puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox"] },
// });

// client.on("qr", (qr) => qrcode.generate(qr, { small: true }));
// client.on("ready", () => console.log("✅ Client is ready!"));

// let isWaitingForSearch = false;
// let isWaitingForUpdate = false;
// let isWaitingForSelection = false;
// let isAddingNewItem = false;
// let isWaitingForValue = false;
// let isInitialMenu = true;
// let currentItem = null;
// let updateField = null;
// let foundItems = [];
// let newItem = {};
// let currentFieldIndex = 0;
// let fieldsToAdd = []; // Initialize fieldsToAdd as an empty array

// client.on("message", async (message) => {
//   try {
//     const userMessage = message.body.trim();

//     if (userMessage.toLowerCase() === "hi") {
//       await loadExcelFile(); // Load the latest file from OneDrive
//       fieldsToAdd = Object.keys(data[0] || {}); // Update fieldsToAdd after loading data
//       message.reply("👋 Hi! Please choose an option:\n1. See/Update details of existing data\n2. Add a new item");
//       isInitialMenu = true;
//     } else if (isAddingNewItem) {
//       const currentField = fieldsToAdd[currentFieldIndex];
//       const input = userMessage.trim().toLowerCase();
//       let value = ["n/a", "na", "n/a", "n/a"].includes(input) ? "" : userMessage;

//       if (currentField.toLowerCase() === "min level" && "Max Level" in newItem) {
//         const maxLevel = newItem["Max Level"];
//         if (!isNaN(value) && !isNaN(maxLevel) && Number(value) > Number(maxLevel)) {
//           message.reply(`❌ Invalid input. Min Level (${value}) cannot be greater than Max Level (${maxLevel}). Please enter a valid value.`);
//           return;
//         }
//       } else if (currentField.toLowerCase() === "max level" && "Min Level" in newItem) {
//         const minLevel = newItem["Min Level"];
//         if (!isNaN(value) && !isNaN(minLevel) && Number(minLevel) > Number(value)) {
//           message.reply(`❌ Invalid input. Max Level (${value}) cannot be less than Min Level (${minLevel}). Please enter a valid value.`);
//           return;
//         }
//       }

//       newItem[currentField] = value;
//       currentFieldIndex++;

//       if (currentFieldIndex < fieldsToAdd.length) {
//         message.reply(`✏️ Enter ${fieldsToAdd[currentFieldIndex]}:`);
//       } else {
//         data.push(newItem);
//         await updateExcelFile(); // Update and upload the file
//         message.reply("✅ New item added and Excel file updated on OneDrive!");
//         isAddingNewItem = false;
//         newItem = {};
//         currentFieldIndex = 0;
//       }
//     } else if (isWaitingForSearch) {
//       const input = userMessage;
//       foundItems = data.filter(
//         (entry) =>
//           (entry["Material Code"] != null &&
//             entry["Material Code"].toString().trim().toLowerCase().startsWith(input.toLowerCase())) ||
//           (typeof entry["Item Description"] === "string" &&
//             entry["Item Description"].toLowerCase().includes(input.toLowerCase()))
//       );

//       if (foundItems.length === 1) {
//         currentItem = foundItems[0];
//         let replyMessage = "📦 Item Details:\n";
//         Object.keys(data[0]).forEach((key) => {
//           replyMessage += `📌 ${key}: ${currentItem[key] || "N/A"}\n`;
//         });
//         replyMessage += "\n⚙️ Do you want to update any information? (yes/no)";
//         message.reply(replyMessage);
//         isWaitingForSearch = false;
//         isWaitingForUpdate = true;
//       } else if (foundItems.length > 1) {
//         let replyMessage = "🔍 Multiple items found. Please select one:\n";
//         foundItems.forEach((item, index) => {
//           replyMessage += `${index + 1}. ${item["Material Code"]} - ${item["Item Description"]}\n`;
//         });
//         replyMessage += "\nReply with the number of the item you want to select.";
//         message.reply(replyMessage);
//         isWaitingForSearch = false;
//         isWaitingForSelection = true;
//       } else {
//         message.reply("❌ No matching Material Code or Item Description found. Please try again.");
//       }
//     } else if (isInitialMenu && userMessage === "1") {
//       isWaitingForSearch = true;
//       isAddingNewItem = false;
//       isInitialMenu = false;
//       message.reply("🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to check details.");
//     } else if (isInitialMenu && userMessage === "2") {
//       isAddingNewItem = true;
//       isWaitingForSearch = false;
//       isInitialMenu = false;
//       currentFieldIndex = 0;
//       message.reply(`✏️ Enter ${fieldsToAdd[currentFieldIndex]}:`);
//     } else if (isWaitingForSelection) {
//       const selectedIndex = parseInt(userMessage) - 1;
//       if (!isNaN(selectedIndex) && selectedIndex >= 0 && selectedIndex < foundItems.length) {
//         currentItem = foundItems[selectedIndex];
//         let replyMessage = "📋 *Item Details:*\n";
//         for (const key in currentItem) {
//           replyMessage += `📌 *${key}:* ${currentItem[key]}\n`;
//         }
//         replyMessage += "\n❓ Do you want to update any information? (yes/no)";
//         message.reply(replyMessage);
//         isWaitingForSelection = false;
//         isWaitingForUpdate = true;
//       } else {
//         message.reply("⚠️ Invalid selection. Please try again.");
//       }
//     } else if (isWaitingForUpdate) {
//       if (isWaitingForValue) {
//         const newValue = userMessage;
//         const fieldName = updateField;

//         if (fieldName.toLowerCase() === "min level") {
//           const maxLevelField = Object.keys(currentItem).find((key) => key.toLowerCase() === "max level");
//           const maxLevel = currentItem[maxLevelField];
//           if (maxLevelField && !isNaN(maxLevel) && !isNaN(newValue) && Number(newValue) > Number(maxLevel)) {
//             message.reply(`❌ Invalid input. Min Level cannot be greater than Max Level (${maxLevel}). Please enter a valid value.`);
//             return;
//           }
//         }
//         currentItem[updateField] = newValue;
//         await updateExcelFile(); // Update and upload the file
//         message.reply(`✅ ${updateField} updated successfully! New value: ${currentItem[updateField]}`);
//         message.reply("⚙️ Do you want to update anything else? (yes/no)");
//         isWaitingForValue = false;
//         updateField = null;
//       } else if (userMessage.toLowerCase() === "yes") {
//         let availableFields = Object.keys(currentItem).map((field) => `📌 ${field}`).join("\n");
//         message.reply(`🛠️ Which field would you like to update?\nAvailable fields:\n${availableFields}`);
//         isWaitingForUpdate = true;
//       } else if (userMessage.toLowerCase() === "no") {
//         message.reply("👍 Thank you! If you need further assistance, say 'hi'.");
//         isWaitingForUpdate = false;
//         currentItem = null;
//       } else if (Object.keys(currentItem).map((key) => key.toLowerCase()).includes(userMessage.toLowerCase())) {
//         updateField = Object.keys(currentItem).find((key) => key.toLowerCase() === userMessage.toLowerCase());
//         isWaitingForValue = true;
//         message.reply(`✏️ Enter the new value for ${updateField}:`);
//       } else {
//         message.reply("❌ Invalid option. Please choose a valid field to update.");
//       }
//     }
//   } catch (error) {
//     console.error("Error handling message:", error);
//     message.reply("⚠️ An error occurred. Please try again.");
//   }
// });

// client.initialize();





































// const { Client, LocalAuth } = require("whatsapp-web.js");
// const qrcode = require("qrcode-terminal");
// const xlsx = require("xlsx");
// const axios = require("axios");
// const fs = require("fs");
// const { exec } = require("child_process");


// // Shared OneDrive link (with edit permissions)
// const oneDriveSharedLink = "https://jindalstainless0-my.sharepoint.com/personal/lucky1_jindalstainless_com/_layouts/15/download.aspx?share=EQe6CWtVC1RFv9NClNFSsm8Btv3U7q3o45cpQorp5-QYsg"; 
// // Replace with your shared link

// // Path to the local Excel file
// const filePath = "./a_data.xlsx";

// // Global variables
// let workbook;
// let sheetName;
// let worksheet;
// let headerRow;
// let data = []; // Initialize data as an empty array

// // Function to download the file from OneDrive

// async function downloadFileFromOneDrive() {
//   try {
//     const response = await axios.get(oneDriveSharedLink, { responseType: "arraybuffer" });

//     // Save the file locally
//     fs.writeFileSync(filePath, response.data);
//     console.log("File downloaded from OneDrive.");
//   } catch (error) {
//     console.error("Error downloading file:", error.message);

//     // If downloading fails, create an empty Excel file
//     if (!fs.existsSync(filePath)) {
//       console.log("Creating an empty Excel file.");
//       workbook = xlsx.utils.book_new();
//       xlsx.utils.book_append_sheet(workbook, xlsx.utils.aoa_to_sheet([[]]), "Sheet1");
//       xlsx.writeFile(workbook, filePath);
//     }
//   }
// }



// async function uploadFileToOneDrive() {
//   return new Promise((resolve, reject) => {
//     // Run the Python script
//     const pythonCommand = `python3 ./upload_to_onedrive_selenium.py`;

//     exec(pythonCommand, (error, stdout, stderr) => {
//       if (error) {
//         console.error("Error uploading file:", stderr);
//         reject(stderr);
//       } else {
//         console.log("File uploaded to OneDrive:", stdout);
//         resolve(stdout);
//       }
//     });
//   });
// }
// // Function to load the Excel file
// async function loadExcelFile() {
//   try {
//     // Download the file from OneDrive
//     await downloadFileFromOneDrive();

//     // Check if the file exists locally
//     if (!fs.existsSync(filePath)) {
//       console.error("File not found locally. Creating an empty Excel file.");
//       workbook = xlsx.utils.book_new(); // Create a new workbook
//       xlsx.utils.book_append_sheet(workbook, xlsx.utils.aoa_to_sheet([[]]), "Sheet1"); // Add an empty sheet
//       xlsx.writeFile(workbook, filePath); // Save the empty workbook
//     }

//     // Load the Excel file
//     workbook = xlsx.readFile(filePath);
//     sheetName = workbook.SheetNames[0];
//     worksheet = workbook.Sheets[sheetName];
//     headerRow = xlsx.utils.sheet_to_json(worksheet, { header: 1 })[0]; // Get the header row

//     data = xlsx.utils.sheet_to_json(worksheet);

//     // Normalize data fields, ensuring all header fields are present
//     data = data.map((entry) => {
//       let normalizedEntry = {};
//       headerRow.forEach((header) => {
//         const trimmedHeader = header?.toString().trim(); // Handle potential undefined header
//         normalizedEntry[trimmedHeader] = entry[trimmedHeader] || "N/A"; // Use value or "N/A"
//       });
//       return normalizedEntry;
//     });

//     // console.log("📦 Loaded data from local file:", JSON.stringify(data, null, 2));
//   } catch (error) {
//     console.error("Error loading Excel file:", error.message);
//   }
// }

// // Function to update Excel file
// async function updateExcelFile() {
//   try {
//     const updatedWorksheet = xlsx.utils.json_to_sheet(data);
//     workbook.Sheets[sheetName] = updatedWorksheet;
//     xlsx.writeFile(workbook, filePath);
//     console.log("Excel file updated locally.");
//     await uploadFileToOneDrive();
//   } catch (error) {
//     console.error("Error updating Excel file:", error);
//   }
// }

// // Initialize WhatsApp client
// const client = new Client({
//   authStrategy: new LocalAuth(),
//   puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox"] },
// });

// client.on("qr", (qr) => qrcode.generate(qr, { small: true }));
// client.on("ready", () => console.log("✅ Client is ready!"));

// let isWaitingForSearch = false;
// let isWaitingForUpdate = false;
// let isWaitingForSelection = false;
// let isAddingNewItem = false;
// let isWaitingForValue = false;
// let isInitialMenu = true;
// let currentItem = null;
// let updateField = null;
// let foundItems = [];
// let newItem = {};
// let currentFieldIndex = 0;
// let fieldsToAdd = []; // Initialize fieldsToAdd as an empty array

// client.on("message", async (message) => {
//   try {
//     const userMessage = message.body.trim();

//     if (userMessage.toLowerCase() === "hi") {
//       await loadExcelFile(); // Load the latest file from OneDrive
//       fieldsToAdd = Object.keys(data[0] || {}); // Update fieldsToAdd after loading data
//       message.reply("👋 Hi! Please choose an option:\n1. See/Update details of existing data\n2. Add a new item");
//       isInitialMenu = true;
//     } else if (isAddingNewItem) {
//       const currentField = fieldsToAdd[currentFieldIndex];
//       const input = userMessage.trim().toLowerCase();
//       let value = ["n/a", "na", "n/a", "n/a"].includes(input) ? "" : userMessage;

//       if (currentField.toLowerCase() === "min level" && "Max Level" in newItem) {
//         const maxLevel = newItem["Max Level"];
//         if (!isNaN(value) && !isNaN(maxLevel) && Number(value) > Number(maxLevel)) {
//           message.reply(`❌ Invalid input. Min Level (${value}) cannot be greater than Max Level (${maxLevel}). Please enter a valid value.`);
//           return;
//         }
//       } else if (currentField.toLowerCase() === "max level" && "Min Level" in newItem) {
//         const minLevel = newItem["Min Level"];
//         if (!isNaN(value) && !isNaN(minLevel) && Number(minLevel) > Number(value)) {
//           message.reply(`❌ Invalid input. Max Level (${value}) cannot be less than Min Level (${minLevel}). Please enter a valid value.`);
//           return;
//         }
//       }

//       newItem[currentField] = value;
//       currentFieldIndex++;

//       if (currentFieldIndex < fieldsToAdd.length) {
//         message.reply(`✏️ Enter ${fieldsToAdd[currentFieldIndex]}:`);
//       } else {
//         data.push(newItem);
//         await updateExcelFile(); // Update and upload the file
//         message.reply("✅ New item added and Excel file updated on OneDrive!");
//         isAddingNewItem = false;
//         newItem = {};
//         currentFieldIndex = 0;
//       }
//     } else if (isWaitingForSearch) {
//       const input = userMessage;
//       foundItems = data.filter(
//         (entry) =>
//           (entry["Material Code"] != null &&
//             entry["Material Code"].toString().trim().toLowerCase().startsWith(input.toLowerCase())) ||
//           (typeof entry["Item Description"] === "string" &&
//             entry["Item Description"].toLowerCase().includes(input.toLowerCase()))
//       );

//       if (foundItems.length === 1) {
//         currentItem = foundItems[0];
//         let replyMessage = "📦 Item Details:\n";
//         Object.keys(data[0]).forEach((key) => {
//           replyMessage += `📌 ${key}: ${currentItem[key] || "N/A"}\n`;
//         });
//         replyMessage += "\n⚙️ Do you want to update any information? (yes/no)";
//         message.reply(replyMessage);
//         isWaitingForSearch = false;
//         isWaitingForUpdate = true;
//       } else if (foundItems.length > 1) {
//         let replyMessage = "🔍 Multiple items found. Please select one:\n";
//         foundItems.forEach((item, index) => {
//           replyMessage += `${index + 1}. ${item["Material Code"]} - ${item["Item Description"]}\n`;
//         });
//         replyMessage += "\nReply with the number of the item you want to select.";
//         message.reply(replyMessage);
//         isWaitingForSearch = false;
//         isWaitingForSelection = true;
//       } else {
//         message.reply("❌ No matching Material Code or Item Description found. Please try again.");
//       }
//     } else if (isInitialMenu && userMessage === "1") {
//       isWaitingForSearch = true;
//       isAddingNewItem = false;
//       isInitialMenu = false;
//       message.reply("🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to check details.");
//     } else if (isInitialMenu && userMessage === "2") {
//       isAddingNewItem = true;
//       isWaitingForSearch = false;
//       isInitialMenu = false;
//       currentFieldIndex = 0;
//       message.reply(`✏️ Enter ${fieldsToAdd[currentFieldIndex]}:`);
//     } else if (isWaitingForSelection) {
//       const selectedIndex = parseInt(userMessage) - 1;
//       if (!isNaN(selectedIndex) && selectedIndex >= 0 && selectedIndex < foundItems.length) {
//         currentItem = foundItems[selectedIndex];
//         let replyMessage = "📋 *Item Details:*\n";
//         for (const key in currentItem) {
//           replyMessage += `📌 *${key}:* ${currentItem[key]}\n`;
//         }
//         replyMessage += "\n❓ Do you want to update any information? (yes/no)";
//         message.reply(replyMessage);
//         isWaitingForSelection = false;
//         isWaitingForUpdate = true;
//       } else {
//         message.reply("⚠️ Invalid selection. Please try again.");
//       }
//     } else if (isWaitingForUpdate) {
//       if (isWaitingForValue) {
//         const newValue = userMessage;
//         const fieldName = updateField;

//         if (fieldName.toLowerCase() === "min level") {
//           const maxLevelField = Object.keys(currentItem).find((key) => key.toLowerCase() === "max level");
//           const maxLevel = currentItem[maxLevelField];
//           if (maxLevelField && !isNaN(maxLevel) && !isNaN(newValue) && Number(newValue) > Number(maxLevel)) {
//             message.reply(`❌ Invalid input. Min Level cannot be greater than Max Level (${maxLevel}). Please enter a valid value.`);
//             return;
//           }
//         }
//         currentItem[updateField] = newValue;
//         await updateExcelFile(); // Update and upload the file
//         message.reply(`✅ ${updateField} updated successfully! New value: ${currentItem[updateField]}`);
//         message.reply("⚙️ Do you want to update anything else? (yes/no)");
//         isWaitingForValue = false;
//         updateField = null;
//       } else if (userMessage.toLowerCase() === "yes") {
//         let availableFields = Object.keys(currentItem).map((field) => `📌 ${field}`).join("\n");
//         message.reply(`🛠️ Which field would you like to update?\nAvailable fields:\n${availableFields}`);
//         isWaitingForUpdate = true;
//       } else if (userMessage.toLowerCase() === "no") {
//         message.reply("👍 Thank you! If you need further assistance, say 'hi'.");
//         isWaitingForUpdate = false;
//         currentItem = null;
//       } else if (Object.keys(currentItem).map((key) => key.toLowerCase()).includes(userMessage.toLowerCase())) {
//         updateField = Object.keys(currentItem).find((key) => key.toLowerCase() === userMessage.toLowerCase());
//         isWaitingForValue = true;
//         message.reply(`✏️ Enter the new value for ${updateField}:`);
//       } else {
//         message.reply("❌ Invalid option. Please choose a valid field to update.");
//       }
//     }
//   } catch (error) {
//     console.error("Error handling message:", error);
//     message.reply("⚠️ An error occurred. Please try again.");
//   }
// });

// client.initialize();




// const { Client, LocalAuth } = require("whatsapp-web.js");
// const qrcode = require("qrcode-terminal");
// const xlsx = require("xlsx");
// const fs = require("fs");
// const { google } = require("googleapis");
// require("dotenv").config();

// // Google Drive Configuration
// const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
// const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
// const REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN;
// const FILE_ID = process.env.GOOGLE_DRIVE_FILE_ID; // ID of your Excel file in Google Drive
// const FILE_PATH = "./a_data.xlsx";

// const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, "urn:ietf:wg:oauth:2.0:oob");
// oauth2Client.setCredentials({ refresh_token: REFRESH_TOKEN });
// const drive = google.drive({ version: "v3", auth: oauth2Client });

// // Global variables
// let workbook;
// let sheetName;
// let worksheet;
// let headerRow;
// let data = []; // Initialize data as an empty array

// // Function to download the file from Google Drive
// async function downloadFileFromGoogleDrive() {
//   try {
//     const response = await drive.files.get(
//       { fileId: FILE_ID, alt: "media" },
//       { responseType: "stream" }
//     );
//     const writer = fs.createWriteStream(FILE_PATH);
//     response.data.pipe(writer);
//     return new Promise((resolve, reject) => {
//       writer.on("finish", () => {
//         console.log("File downloaded from Google Drive.");
//         resolve();
//       });
//       writer.on("error", (err) => {
//         console.error("Error downloading file:", err.message);
//         reject(err);
//       });
//     });
//   } catch (error) {
//     console.error("Error downloading file from Google Drive:", error.message);
//     throw error;
//   }
// }

// async function listFiles() {
//   try {
//     const response = await drive.files.list({
//       q: "name='a_data.xlsx'",
//       fields: "files(id, name, permissions, parents)",
//       spaces: "drive", // Include both My Drive and Shared Drives
//       corpora: "user", // Search all files the user has access to
//       includeItemsFromAllDrives: true, // Include Shared Drives
//       supportsAllDrives: true, // Support Shared Drives
//     });
//     console.log("Files found:", response.data.files);
//     if (response.data.files.length === 0) {
//       console.log("No files found with name 'a_data.xlsx'. Checking case-insensitive...");
//       const responseCaseInsensitive = await drive.files.list({
//         q: "name contains 'a_data.xlsx'",
//         fields: "files(id, name, permissions, parents)",
//         spaces: "drive",
//         corpora: "user",
//         includeItemsFromAllDrives: true,
//         supportsAllDrives: true,
//       });
//       console.log("Files found (case-insensitive):", responseCaseInsensitive.data.files);
//     }
//   } catch (error) {
//     console.error("Error listing files:", error.response ? error.response.data : error.message);
//   }
// }
// listFiles();

// // Function to upload the file to Google Drive
// async function uploadFileToGoogleDrive() {
//   try {
//     const fileContent = fs.createReadStream(FILE_PATH);
//     await drive.files.update({
//       fileId: FILE_ID,
//       media: {
//         mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
//         body: fileContent,
//       },
//     });
//     console.log("File uploaded to Google Drive.");
//   } catch (error) {
//     console.error("Error uploading file to Google Drive:", error.message);
//     throw error;
//   }
// }

// // Function to load the Excel file
// async function loadExcelFile() {
//   try {
//     await downloadFileFromGoogleDrive();

//     if (!fs.existsSync(FILE_PATH)) {
//       console.error("File not found locally. Creating an empty Excel file.");
//       workbook = xlsx.utils.book_new();
//       xlsx.utils.book_append_sheet(workbook, xlsx.utils.aoa_to_sheet([[]]), "Sheet1");
//       xlsx.writeFile(workbook, FILE_PATH);
//     }

//     workbook = xlsx.readFile(FILE_PATH);
//     sheetName = workbook.SheetNames[0];
//     worksheet = workbook.Sheets[sheetName];
//     headerRow = xlsx.utils.sheet_to_json(worksheet, { header: 1 })[0];

//     data = xlsx.utils.sheet_to_json(worksheet);
//     data = data.map((entry) => {
//       let normalizedEntry = {};
//       headerRow.forEach((header) => {
//         const trimmedHeader = header?.toString().trim();
//         normalizedEntry[trimmedHeader] = entry[trimmedHeader] || "N/A";
//       });
//       return normalizedEntry;
//     });

//     console.log("📦 Loaded data from local file:", JSON.stringify(data, null, 2));
//   } catch (error) {
//     console.error("Error loading Excel file:", error.message);
//   }
// }

// // Function to update Excel file
// async function updateExcelFile() {
//   try {
//     const updatedWorksheet = xlsx.utils.json_to_sheet(data);
//     workbook.Sheets[sheetName] = updatedWorksheet;
//     xlsx.writeFile(workbook, FILE_PATH);
//     console.log("Excel file updated locally.");
//     await uploadFileToGoogleDrive();
//   } catch (error) {
//     console.error("Error updating Excel file:", error);
//   }
// }

// // Initialize WhatsApp client
// const client = new Client({
//   authStrategy: new LocalAuth(),
//   puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox"] },
// });

// client.on("qr", (qr) => qrcode.generate(qr, { small: true }));
// client.on("ready", () => console.log("✅ Client is ready!"));

// let isWaitingForSearch = false;
// let isWaitingForUpdate = false;
// let isWaitingForSelection = false;
// let isAddingNewItem = false;
// let isWaitingForValue = false;
// let isInitialMenu = true;
// let currentItem = null;
// let updateField = null;
// let foundItems = [];
// let newItem = {};
// let currentFieldIndex = 0;
// let fieldsToAdd = [];

// client.on("message", async (message) => {
//   try {
//     const userMessage = message.body.trim();

//     if (userMessage.toLowerCase() === "hi") {
//       await loadExcelFile();
//       fieldsToAdd = Object.keys(data[0] || {});
//       message.reply("👋 Hi! Please choose an option:\n1. See/Update details of existing data\n2. Add a new item");
//       isInitialMenu = true;
//     } else if (isAddingNewItem) {
//       const currentField = fieldsToAdd[currentFieldIndex];
//       const input = userMessage.trim().toLowerCase();
//       let value = ["n/a", "na", "n/a", "n/a"].includes(input) ? "" : userMessage;

//       if (currentField.toLowerCase() === "min level" && "Max Level" in newItem) {
//         const maxLevel = newItem["Max Level"];
//         if (!isNaN(value) && !isNaN(maxLevel) && Number(value) > Number(maxLevel)) {
//           message.reply(`❌ Invalid input. Min Level (${value}) cannot be greater than Max Level (${maxLevel}). Please enter a valid value.`);
//           return;
//         }
//       } else if (currentField.toLowerCase() === "max level" && "Min Level" in newItem) {
//         const minLevel = newItem["Min Level"];
//         if (!isNaN(value) && !isNaN(minLevel) && Number(minLevel) > Number(value)) {
//           message.reply(`❌ Invalid input. Max Level (${value}) cannot be less than Min Level (${minLevel}). Please enter a valid value.`);
//           return;
//         }
//       }

//       newItem[currentField] = value;
//       currentFieldIndex++;

//       if (currentFieldIndex < fieldsToAdd.length) {
//         message.reply(`✏️ Enter ${fieldsToAdd[currentFieldIndex]}:`);
//       } else {
//         data.push(newItem);
//         await updateExcelFile();
//         message.reply("✅ New item added and Excel file updated on Google Drive!");
//         isAddingNewItem = false;
//         newItem = {};
//         currentFieldIndex = 0;
//       }
//     } else if (isWaitingForSearch) {
//       const input = userMessage;
//       foundItems = data.filter(
//         (entry) =>
//           (entry["Material Code"] != null &&
//             entry["Material Code"].toString().trim().toLowerCase().startsWith(input.toLowerCase())) ||
//           (typeof entry["Item Description"] === "string" &&
//             entry["Item Description"].toLowerCase().includes(input.toLowerCase()))
//       );

//       if (foundItems.length === 1) {
//         currentItem = foundItems[0];
//         let replyMessage = "📦 Item Details:\n";
//         Object.keys(data[0]).forEach((key) => {
//           replyMessage += `📌 ${key}: ${currentItem[key] || "N/A"}\n`;
//         });
//         replyMessage += "\n⚙️ Do you want to update any information? (yes/no)";
//         message.reply(replyMessage);
//         isWaitingForSearch = false;
//         isWaitingForUpdate = true;
//       } else if (foundItems.length > 1) {
//         let replyMessage = "🔍 Multiple items found. Please select one:\n";
//         foundItems.forEach((item, index) => {
//           replyMessage += `${index + 1}. ${item["Material Code"]} - ${item["Item Description"]}\n`;
//         });
//         replyMessage += "\nReply with the number of the item you want to select.";
//         message.reply(replyMessage);
//         isWaitingForSearch = false;
//         isWaitingForSelection = true;
//       } else {
//         message.reply("❌ No matching Material Code or Item Description found. Please try again.");
//       }
//     } else if (isInitialMenu && userMessage === "1") {
//       isWaitingForSearch = true;
//       isAddingNewItem = false;
//       isInitialMenu = false;
//       message.reply("🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to check details.");
//     } else if (isInitialMenu && userMessage === "2") {
//       isAddingNewItem = true;
//       isWaitingForSearch = false;
//       isInitialMenu = false;
//       currentFieldIndex = 0;
//       message.reply(`✏️ Enter ${fieldsToAdd[currentFieldIndex]}:`);
//     } else if (isWaitingForSelection) {
//       const selectedIndex = parseInt(userMessage) - 1;
//       if (!isNaN(selectedIndex) && selectedIndex >= 0 && selectedIndex < foundItems.length) {
//         currentItem = foundItems[selectedIndex];
//         let replyMessage = "📋 *Item Details:*\n";
//         for (const key in currentItem) {
//           replyMessage += `📌 *${key}:* ${currentItem[key]}\n`;
//         }
//         replyMessage += "\n❓ Do you want to update any information? (yes/no)";
//         message.reply(replyMessage);
//         isWaitingForSelection = false;
//         isWaitingForUpdate = true;
//       } else {
//         message.reply("⚠️ Invalid selection. Please try again.");
//       }
//     } else if (isWaitingForUpdate) {
//       if (isWaitingForValue) {
//         const newValue = userMessage;
//         const fieldName = updateField;

//         if (fieldName.toLowerCase() === "min level") {
//           const maxLevelField = Object.keys(currentItem).find((key) => key.toLowerCase() === "max level");
//           const maxLevel = currentItem[maxLevelField];
//           if (maxLevelField && !isNaN(maxLevel) && !isNaN(newValue) && Number(newValue) > Number(maxLevel)) {
//             message.reply(`❌ Invalid input. Min Level cannot be greater than Max Level (${maxLevel}). Please enter a valid value.`);
//             return;
//           }
//         }
//         currentItem[updateField] = newValue;
//         await updateExcelFile();
//         message.reply(`✅ ${updateField} updated successfully! New value: ${currentItem[updateField]}`);
//         message.reply("⚙️ Do you want to update anything else? (yes/no)");
//         isWaitingForValue = false;
//         updateField = null;
//       } else if (userMessage.toLowerCase() === "yes") {
//         let availableFields = Object.keys(currentItem).map((field) => `📌 ${field}`).join("\n");
//         message.reply(`🛠️ Which field would you like to update?\nAvailable fields:\n${availableFields}`);
//         isWaitingForUpdate = true;
//       } else if (userMessage.toLowerCase() === "no") {
//         message.reply("👍 Thank you! If you need further assistance, say 'hi'.");
//         isWaitingForUpdate = false;
//         currentItem = null;
//       } else if (Object.keys(currentItem).map((key) => key.toLowerCase()).includes(userMessage.toLowerCase())) {
//         updateField = Object.keys(currentItem).find((key) => key.toLowerCase() === userMessage.toLowerCase());
//         isWaitingForValue = true;
//         message.reply(`✏️ Enter the new value for ${updateField}:`);
//       } else {
//         message.reply("❌ Invalid option. Please choose a valid field to update.");
//       }
//     }
//   } catch (error) {
//     console.error("Error handling message:", error);
//     message.reply("⚠️ An error occurred. Please try again.");
//   }
// });

// client.initialize();











// const { Client, LocalAuth } = require("whatsapp-web.js");
// const qrcode = require("qrcode-terminal");
// const xlsx = require("xlsx");
// const fs = require("fs");
// const { google } = require("googleapis");
// require("dotenv").config();

// // Constants
// const FILE_PATH = "./a_data.xlsx";
// const SHEET_NAME_DEFAULT = "Sheet1";
// const SERIAL_NUMBER_FIELD = "S.No"; // The field name for the serial number in the Excel file
// const REQUIRED_FIELDS = ["Material Code", "Item Description"]; // Required fields for adding a new item
// const MESSAGES = {
//   WELCOME: "👋 Hi! Please choose an option:\n1. See/Update details of existing data\n2. Add a new item\n3. List all items\n4. Delete an item\nType 'help' for more commands.",
//   SEARCH_PROMPT: "🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to check details (or type 'cancel' to go back):",
//   DELETE_PROMPT: "🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to delete (or type 'cancel' to go back):",
//   ADD_PROMPT: (field) => `✏️ Enter ${field} (or type 'cancel' to go back):`,
//   ADD_CONFIRM: "✅ Material Code and Item Description added. Type 'done' to finish (remaining fields will be set to N/A) or 'continue' to add more details:",
//   UPDATE_PROMPT: (fields) => `🛠️ Which field would you like to update?\nAvailable fields:\n${fields}\n(or type 'cancel' to go back)`,
//   UPDATE_VALUE_PROMPT: (field) => `✏️ Enter the new value for ${field} (or type 'cancel' to go back):`,
//   SELECT_ITEM_PROMPT: (items) => `🔍 Multiple items found. Please select one:\n${items}\nReply with the number of the item (or type 'cancel' to go back).`,
//   ITEM_DETAILS: (item) => {
//     let reply = "📋 *Item Details:*\n";
//     for (const key in item) {
//       reply += `📌 *${key}:* ${item[key]}\n`;
//     }
//     return reply + "\n❓ Do you want to update any information? (yes/no)";
//   },
//   SUCCESS_ADD: "✅ New item added and Excel file updated on Google Drive!",
//   SUCCESS_DELETE: "✅ Item deleted successfully and Excel file updated on Google Drive!",
//   SUCCESS_UPDATE: (field, value) => `✅ ${field} updated successfully! New value: ${value}`,
//   UPDATE_CONFIRM: "⚙️ Do you want to update anything else? (yes/no)",
//   ERROR: "⚠️ An error occurred. Please try again or type 'hi' to restart.",
//   INVALID_INPUT: "❌ Invalid input. Please try again.",
//   NO_ITEMS_FOUND: "❌ No matching Material Code or Item Description found. Please try again.",
//   INVALID_SELECTION: "⚠️ Invalid selection. Please try again.",
//   CANCEL: "👍 Operation cancelled. Type 'hi' to start again.",
//   HELP: "📖 Available commands:\n- 'hi': Start the bot\n- 'help': Show this message\n- 'cancel': Cancel the current operation\nChoose an option:\n1. See/Update details\n2. Add a new item\n3. List all items\n4. Delete an item",
//   LOADING: "⏳ Loading data, please wait...",
//   LIST_ITEMS: (items) => `📋 *All Items:*\n${items}\nType 'hi' to go back to the main menu.`,
//   MISSING_REQUIRED_FIELDS: "❌ Both Material Code and Item Description are required. Item not added. Type 'hi' to start again.",
//   INVALID_MIN_LEVEL: (maxLevel) => `❌ Invalid input. Min Level must be less than or equal to Max Level (${maxLevel}). Please enter a valid value.`,
//   INVALID_MAX_LEVEL: (minLevel) => `❌ Invalid input. Max Level must be greater than or equal to Min Level (${minLevel}). Please enter a valid value.`,
//   INVALID_NUMERIC: "❌ Min Level and Max Level must be numeric values. Please enter a valid number.",
//   DUPLICATE_MATERIAL_CODE_PROMPT: (item) =>
//     `⚠️ The Material Code already exists:\n${item["S.No"]} - ${item["Material Code"]} - ${item["Item Description"]}\nWould you like to update this item? (yes/no)`,
// };

// // Google Drive Configuration
// const oauth2Client = new google.auth.OAuth2(
//   process.env.GOOGLE_CLIENT_ID,
//   process.env.GOOGLE_CLIENT_SECRET,
//   "urn:ietf:wg:oauth:2.0:oob"
// );
// oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
// const drive = google.drive({ version: "v3", auth: oauth2Client });

// // State Management
// const BotState = {
//   IDLE: "IDLE",
//   INITIAL_MENU: "INITIAL_MENU",
//   SEARCHING: "SEARCHING",
//   SELECTING_ITEM: "SELECTING_ITEM",
//   UPDATING: "UPDATING",
//   UPDATING_VALUE: "UPDATING_VALUE",
//   ADDING_ITEM: "ADDING_ITEM",
//   ADDING_CONFIRM: "ADDING_CONFIRM",
//   LISTING_ITEMS: "LISTING_ITEMS",
//   DELETING_ITEM: "DELETING_ITEM",
//   SELECTING_ITEM_TO_DELETE: "SELECTING_ITEM_TO_DELETE",
//   HANDLING_DUPLICATE: "HANDLING_DUPLICATE", // New state for handling duplicate Material Codes
// };

// // Class to manage bot state and data
// class InventoryBot {
//   constructor() {
//     this.state = BotState.IDLE;
//     this.data = [];
//     this.workbook = null;
//     this.sheetName = null;
//     this.worksheet = null;
//     this.headerRow = null;
//     this.currentItem = null;
//     this.foundItems = [];
//     this.newItem = {};
//     this.currentFieldIndex = 0;
//     this.fieldsToAdd = [];
//     this.requiredFieldsEntered = 0;
//     this.updateField = null;
//     this.client = new Client({
//       authStrategy: new LocalAuth(),
//       puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox"] },
//     });

//     // Bind methods
//     this.handleMessage = this.handleMessage.bind(this);
//     this.downloadFileFromGoogleDrive = this.downloadFileFromGoogleDrive.bind(this);
//     this.uploadFileToGoogleDrive = this.uploadFileToGoogleDrive.bind(this);
//     this.loadExcelFile = this.loadExcelFile.bind(this);
//     this.updateExcelFile = this.updateExcelFile.bind(this);
//     this.listFiles = this.listFiles.bind(this);
//     this.generateSerialNumber = this.generateSerialNumber.bind(this);

//     // Initialize WhatsApp client
//     this.client.on("qr", (qr) => qrcode.generate(qr, { small: true }));
//     this.client.on("ready", () => console.log("✅ Client is ready!"));
//     this.client.on("message", this.handleMessage);
//     this.client.initialize();
//   }

//   // Generate the next serial number
//   generateSerialNumber() {
//     if (this.data.length === 0) {
//       return 1; // If the Excel file is empty, start with S.No 1
//     }
//     const serialNumbers = this.data
//       .map((item) => parseInt(item[SERIAL_NUMBER_FIELD], 10))
//       .filter((num) => !isNaN(num));
//     const maxSerialNumber = Math.max(...serialNumbers, 0);
//     return maxSerialNumber + 1;
//   }

//   // Google Drive: Download file with retry mechanism
//   async downloadFileFromGoogleDrive() {
//     const maxRetries = 3;
//     let attempt = 1;
//     while (attempt <= maxRetries) {
//       try {
//         console.log(`[DEBUG] Attempt ${attempt} to download file from Google Drive...`);
//         const response = await drive.files.get(
//           { fileId: process.env.GOOGLE_DRIVE_FILE_ID, alt: "media" },
//           { responseType: "stream" }
//         );
//         const writer = fs.createWriteStream(FILE_PATH);
//         response.data.pipe(writer);
//         return new Promise((resolve, reject) => {
//           writer.on("finish", () => {
//             console.log("File downloaded from Google Drive.");
//             resolve();
//           });
//           writer.on("error", (err) => {
//             console.error("Error downloading file:", err.message);
//             reject(err);
//           });
//         });
//       } catch (error) {
//         console.error(`Error downloading file from Google Drive (attempt ${attempt}):`, error.message);
//         if (attempt === maxRetries) throw error;
//         attempt++;
//         await new Promise(resolve => setTimeout(resolve, 2000)); // Wait 2 seconds before retrying
//       }
//     }
//   }

//   // Google Drive: Upload file
//   async uploadFileToGoogleDrive() {
//     try {
//       const fileContent = fs.createReadStream(FILE_PATH);
//       await drive.files.update({
//         fileId: process.env.GOOGLE_DRIVE_FILE_ID,
//         media: {
//           mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
//           body: fileContent,
//         },
//       });
//       console.log("File uploaded to Google Drive.");
//     } catch (error) {
//       console.error("Error uploading file to Google Drive:", error.message);
//       throw error;
//     }
//   }

//   // Google Drive: List files (for debugging)
//   async listFiles() {
//     try {
//       const response = await drive.files.list({
//         q: "name='a_data.xlsx'",
//         fields: "files(id, name, permissions, parents)",
//         spaces: "drive",
//         corpora: "user",
//         includeItemsFromAllDrives: true,
//         supportsAllDrives: true,
//       });
//       console.log("Files found:", response.data.files);
//       if (response.data.files.length === 0) {
//         console.log("No files found with name 'a_data.xlsx'. Checking case-insensitive...");
//         const responseCaseInsensitive = await drive.files.list({
//           q: "name contains 'a_data.xlsx'",
//           fields: "files(id, name, permissions, parents)",
//           spaces: "drive",
//           corpora: "user",
//           includeItemsFromAllDrives: true,
//           supportsAllDrives: true,
//         });
//         console.log("Files found (case-insensitive):", responseCaseInsensitive.data.files);
//       }
//     } catch (error) {
//       console.error("Error listing files:", error.response ? error.response.data : error.message);
//     }
//   }

//   // Excel: Load file
//   async loadExcelFile(message) {
//     try {
//       if (message) message.reply(MESSAGES.LOADING);
//       console.log("[DEBUG] Downloading file from Google Drive...");
//       // Force a fresh download by deleting the local file
//       if (fs.existsSync(FILE_PATH)) {
//         console.log("[DEBUG] Deleting existing local file to force fresh download...");
//         fs.unlinkSync(FILE_PATH);
//       }
//       await this.downloadFileFromGoogleDrive();

//       if (!fs.existsSync(FILE_PATH)) {
//         console.error("File not found locally. Creating an empty Excel file.");
//         this.workbook = xlsx.utils.book_new();
//         xlsx.utils.book_append_sheet(this.workbook, xlsx.utils.aoa_to_sheet([[]]), SHEET_NAME_DEFAULT);
//         xlsx.writeFile(this.workbook, FILE_PATH);
//       }

//       console.log("[DEBUG] Reading local Excel file...");
//       this.workbook = xlsx.readFile(FILE_PATH);
//       this.sheetName = this.workbook.SheetNames[0] || SHEET_NAME_DEFAULT;
//       console.log("[DEBUG] Using sheet:", this.sheetName);
//       this.worksheet = this.workbook.Sheets[this.sheetName];
//       this.headerRow = xlsx.utils.sheet_to_json(this.worksheet, { header: 1 })[0] || [];
//       console.log("[DEBUG] Header row:", this.headerRow);

//       this.data = xlsx.utils.sheet_to_json(this.worksheet);
//       this.data = this.data.map((entry) => {
//         let normalizedEntry = {};
//         this.headerRow.forEach((header) => {
//           const trimmedHeader = header?.toString().trim();
//           normalizedEntry[trimmedHeader] = entry[trimmedHeader] || "N/A";
//         });
//         return normalizedEntry;
//       });

//       console.log("📦 Loaded data from local file:", JSON.stringify(this.data, null, 2));
//     } catch (error) {
//       console.error("Error loading Excel file:", error.message);
//       throw error;
//     }
//   }

//   // Excel: Update file
//   async updateExcelFile() {
//     try {
//       const updatedWorksheet = xlsx.utils.json_to_sheet(this.data);
//       this.workbook.Sheets[this.sheetName] = updatedWorksheet;
//       xlsx.writeFile(this.workbook, FILE_PATH);
//       console.log("Excel file updated locally.");
//       await this.uploadFileToGoogleDrive();
//     } catch (error) {
//       console.error("Error updating Excel file:", error.message);
//       throw error;
//     }
//   }

//   // Fill remaining fields with "N/A"
//   fillRemainingFields() {
//     this.fieldsToAdd.forEach((field) => {
//       if (!(field in this.newItem)) {
//         this.newItem[field] = "N/A";
//       }
//     });
//   }

//   // Validate required fields
//   validateRequiredFields() {
//     return REQUIRED_FIELDS.every((field) => {
//       const value = this.newItem[field];
//       return value && value.trim() !== "" && value !== "N/A";
//     });
//   }

//   // Check if a value is defined (not "N/A" and not empty)
//   isDefined(value) {
//     return value && value !== "N/A" && value.toString().trim() !== "";
//   }

//   // WhatsApp: Handle messages
//   async handleMessage(message) {
//     try {
//       const userMessage = message.body.trim().toLowerCase();
//       console.log(`[DEBUG] Current state: ${this.state}, User message: ${userMessage}`);

//       // Handle "cancel" command
//       if (userMessage === "cancel" && this.state !== BotState.IDLE) {
//         console.log("[DEBUG] Cancel command received.");
//         // If required fields are entered, save the item before canceling
//         if (this.state === BotState.ADDING_ITEM || this.state === BotState.ADDING_CONFIRM) {
//           if (this.requiredFieldsEntered === REQUIRED_FIELDS.length && this.validateRequiredFields()) {
//             console.log("[DEBUG] Required fields entered, saving item before canceling.");
//             this.fillRemainingFields();
//             this.data.push(this.newItem);
//             await this.updateExcelFile();
//             message.reply(MESSAGES.SUCCESS_ADD);
//           } else {
//             console.log("[DEBUG] Required fields not fully entered, discarding item.");
//             message.reply(MESSAGES.CANCEL);
//           }
//         } else {
//           message.reply(MESSAGES.CANCEL);
//         }

//         // Reset state
//         this.state = BotState.IDLE;
//         this.currentItem = null;
//         this.foundItems = [];
//         this.newItem = {};
//         this.currentFieldIndex = 0;
//         this.requiredFieldsEntered = 0;
//         this.updateField = null;
//         return;
//       }

//       // Handle "help" command
//       if (userMessage === "help") {
//         console.log("[DEBUG] Help command received.");
//         message.reply(MESSAGES.HELP);
//         return;
//       }

//       // State machine
//       switch (this.state) {
//         case BotState.IDLE:
//           console.log("[DEBUG] In IDLE state.");
//           if (userMessage === "hi") {
//             await this.loadExcelFile(message);
//             if (this.data.length === 0) {
//               message.reply("⚠️ The Excel file is empty. Please add items using option 2.");
//               return;
//             }
//             // Exclude S.No and required fields from fieldsToAdd
//             this.fieldsToAdd = Object.keys(this.data[0] || {}).filter(
//               (field) =>
//                 field.toLowerCase() !== SERIAL_NUMBER_FIELD.toLowerCase() &&
//                 !REQUIRED_FIELDS.map(f => f.toLowerCase()).includes(field.toLowerCase())
//             );
//             message.reply(MESSAGES.WELCOME);
//             this.state = BotState.INITIAL_MENU;
//           }
//           break;

//         case BotState.INITIAL_MENU:
//           console.log("[DEBUG] In INITIAL_MENU state.");
//           if (userMessage === "1") {
//             this.state = BotState.SEARCHING;
//             message.reply(MESSAGES.SEARCH_PROMPT);
//           } else if (userMessage === "2") {
//             if (this.data.length === 0) {
//               message.reply("⚠️ The Excel file is empty. Please add at least one item first.");
//               return;
//             }
//             this.state = BotState.ADDING_ITEM;
//             this.currentFieldIndex = 0;
//             this.requiredFieldsEntered = 0;
//             this.newItem = {};
//             // Automatically generate S.No and add it to newItem
//             this.newItem[SERIAL_NUMBER_FIELD] = this.generateSerialNumber();
//             message.reply(MESSAGES.ADD_PROMPT(REQUIRED_FIELDS[this.currentFieldIndex]));
//           } else if (userMessage === "3") {
//             this.state = BotState.LISTING_ITEMS;
//             let itemsList = this.data.map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item["Material Code"]} - ${item["Item Description"]}`).join("\n");
//             message.reply(MESSAGES.LIST_ITEMS(itemsList));
//             this.state = BotState.IDLE;
//           } else if (userMessage === "4") {
//             this.state = BotState.DELETING_ITEM;
//             message.reply(MESSAGES.DELETE_PROMPT);
//           } else {
//             message.reply(MESSAGES.INVALID_INPUT);
//           }
//           break;

//         case BotState.SEARCHING:
//           console.log("[DEBUG] In SEARCHING state.");
//           const input = message.body.trim();
//           this.foundItems = this.data.filter(
//             (entry) =>
//               (entry["Material Code"] != null &&
//                 entry["Material Code"].toString().trim().toLowerCase().startsWith(input.toLowerCase())) ||
//               (typeof entry["Item Description"] === "string" &&
//                 entry["Item Description"].toLowerCase().includes(input.toLowerCase()))
//           );

//           if (this.foundItems.length === 1) {
//             this.currentItem = this.foundItems[0];
//             message.reply(MESSAGES.ITEM_DETAILS(this.currentItem));
//             this.state = BotState.UPDATING;
//           } else if (this.foundItems.length > 1) {
//             let itemsList = this.foundItems
//               .map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item["Material Code"]} - ${item["Item Description"]}`)
//               .join("\n");
//             message.reply(MESSAGES.SELECT_ITEM_PROMPT(itemsList));
//             this.state = BotState.SELECTING_ITEM;
//           } else {
//             message.reply(MESSAGES.NO_ITEMS_FOUND);
//           }
//           break;

//         case BotState.SELECTING_ITEM:
//           console.log("[DEBUG] In SELECTING_ITEM state.");
//           const selectedIndex = parseInt(userMessage) - 1;
//           if (!isNaN(selectedIndex) && selectedIndex >= 0 && selectedIndex < this.foundItems.length) {
//             this.currentItem = this.foundItems[selectedIndex];
//             message.reply(MESSAGES.ITEM_DETAILS(this.currentItem));
//             this.state = BotState.UPDATING;
//           } else {
//             message.reply(MESSAGES.INVALID_SELECTION);
//           }
//           break;

//         case BotState.DELETING_ITEM:
//           console.log("[DEBUG] In DELETING_ITEM state.");
//           const deleteInput = message.body.trim();
//           this.foundItems = this.data.filter(
//             (entry) =>
//               (entry["Material Code"] != null &&
//                 entry["Material Code"].toString().trim().toLowerCase().startsWith(deleteInput.toLowerCase())) ||
//               (typeof entry["Item Description"] === "string" &&
//                 entry["Item Description"].toLowerCase().includes(deleteInput.toLowerCase()))
//           );

//           if (this.foundItems.length === 1) {
//             this.currentItem = this.foundItems[0];
//             this.data = this.data.filter(
//               (entry) =>
//                 entry["Material Code"] !== this.currentItem["Material Code"] ||
//                 entry["Item Description"] !== this.currentItem["Item Description"]
//             );
//             await this.updateExcelFile();
//             message.reply(MESSAGES.SUCCESS_DELETE);
//             this.currentItem = null;
//             this.foundItems = [];
//             this.state = BotState.IDLE;
//           } else if (this.foundItems.length > 1) {
//             let itemsList = this.foundItems
//               .map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item["Material Code"]} - ${item["Item Description"]}`)
//               .join("\n");
//             message.reply(MESSAGES.SELECT_ITEM_PROMPT(itemsList));
//             this.state = BotState.SELECTING_ITEM_TO_DELETE;
//           } else {
//             message.reply(MESSAGES.NO_ITEMS_FOUND);
//             this.state = BotState.IDLE;
//           }
//           break;

//         case BotState.SELECTING_ITEM_TO_DELETE:
//           console.log("[DEBUG] In SELECTING_ITEM_TO_DELETE state.");
//           const deleteIndex = parseInt(userMessage) - 1;
//           if (!isNaN(deleteIndex) && deleteIndex >= 0 && deleteIndex < this.foundItems.length) {
//             this.currentItem = this.foundItems[deleteIndex];
//             this.data = this.data.filter(
//               (entry) =>
//                 entry["Material Code"] !== this.currentItem["Material Code"] ||
//                 entry["Item Description"] !== this.currentItem["Item Description"]
//             );
//             await this.updateExcelFile();
//             message.reply(MESSAGES.SUCCESS_DELETE);
//             this.currentItem = null;
//             this.foundItems = [];
//             this.state = BotState.IDLE;
//           } else {
//             message.reply(MESSAGES.INVALID_SELECTION);
//             this.state = BotState.IDLE;
//           }
//           break;

//         case BotState.UPDATING:
//           console.log("[DEBUG] In UPDATING state.");
//           if (userMessage === "yes") {
//             let availableFields = Object.keys(this.currentItem).map((field) => `📌 ${field}`).join("\n");
//             message.reply(MESSAGES.UPDATE_PROMPT(availableFields));
//           } else if (userMessage === "no") {
//             message.reply("👍 Thank you! Type 'hi' to start again.");
//             this.currentItem = null;
//             this.state = BotState.IDLE;
//           } else if (Object.keys(this.currentItem).map((key) => key.toLowerCase()).includes(userMessage)) {
//             this.updateField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === userMessage);
//             this.state = BotState.UPDATING_VALUE;
//             message.reply(MESSAGES.UPDATE_VALUE_PROMPT(this.updateField));
//           } else {
//             message.reply(MESSAGES.INVALID_INPUT);
//           }
//           break;

//         case BotState.UPDATING_VALUE:
//           console.log("[DEBUG] In UPDATING_VALUE state.");
//           const newValue = message.body.trim();
//           const fieldName = this.updateField;

//           // Validation for Min Level
//           if (fieldName.toLowerCase() === "min level") {
//             const maxLevelField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === "max level");
//             const maxLevel = maxLevelField ? this.currentItem[maxLevelField] : null;
//             if (this.isDefined(maxLevel)) {
//               if (isNaN(newValue) || isNaN(maxLevel)) {
//                 message.reply(MESSAGES.INVALID_NUMERIC);
//                 return;
//               }
//               if (Number(newValue) > Number(maxLevel)) {
//                 message.reply(MESSAGES.INVALID_MIN_LEVEL(maxLevel));
//                 return;
//               }
//             }
//           }
//           // Validation for Max Level
//           else if (fieldName.toLowerCase() === "max level") {
//             const minLevelField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === "min level");
//             const minLevel = minLevelField ? this.currentItem[minLevelField] : null;
//             if (this.isDefined(minLevel)) {
//               if (isNaN(newValue) || isNaN(minLevel)) {
//                 message.reply(MESSAGES.INVALID_NUMERIC);
//                 return;
//               }
//               if (Number(newValue) < Number(minLevel)) {
//                 message.reply(MESSAGES.INVALID_MAX_LEVEL(minLevel));
//                 return;
//               }
//             }
//           }

//           console.log(`[DEBUG] Updating ${fieldName} to ${newValue}`);
//           this.currentItem[this.updateField] = newValue;
//           await this.updateExcelFile();
//           message.reply(MESSAGES.SUCCESS_UPDATE(this.updateField, newValue));
//           message.reply(MESSAGES.UPDATE_CONFIRM);
//           this.state = BotState.UPDATING;
//           this.updateField = null;
//           console.log("[DEBUG] Transitioned back to UPDATING state.");
//           break;

//         case BotState.ADDING_ITEM:
//           console.log("[DEBUG] In ADDING_ITEM state.");
//           // Determine if we're collecting required fields or additional fields
//           let currentField;
//           if (this.currentFieldIndex < REQUIRED_FIELDS.length) {
//             currentField = REQUIRED_FIELDS[this.currentFieldIndex];
//           } else {
//             const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
//             currentField = this.fieldsToAdd[additionalFieldIndex];
//           }

//           const inputValue = message.body.trim();
//           let value = ["n/a", "na", "n/a", "n/a"].includes(inputValue.toLowerCase()) ? "" : inputValue;

//           // Check for duplicate Material Code when the user enters it
//           if (currentField === "Material Code") {
//             const existingItem = this.data.find(
//               (entry) =>
//                 entry["Material Code"] &&
//                 entry["Material Code"].toString().trim().toLowerCase() === value.toLowerCase()
//             );
//             if (existingItem) {
//               this.currentItem = existingItem; // Store the existing item
//               message.reply(MESSAGES.DUPLICATE_MATERIAL_CODE_PROMPT(existingItem));
//               this.state = BotState.HANDLING_DUPLICATE;
//               return;
//             }
//           }

//           // Validation for Min Level when adding a new item
//           if (currentField.toLowerCase() === "min level" && "Max Level" in this.newItem) {
//             const maxLevel = this.newItem["Max Level"];
//             if (this.isDefined(maxLevel)) {
//               if (isNaN(value) || isNaN(maxLevel)) {
//                 message.reply(MESSAGES.INVALID_NUMERIC);
//                 return;
//               }
//               if (Number(value) > Number(maxLevel)) {
//                 message.reply(MESSAGES.INVALID_MIN_LEVEL(maxLevel));
//                 return;
//               }
//             }
//           }
//           // Validation for Max Level when adding a new item
//           else if (currentField.toLowerCase() === "max level" && "Min Level" in this.newItem) {
//             const minLevel = this.newItem["Min Level"];
//             if (this.isDefined(minLevel)) {
//               if (isNaN(value) || isNaN(minLevel)) {
//                 message.reply(MESSAGES.INVALID_NUMERIC);
//                 return;
//               }
//               if (Number(value) < Number(minLevel)) {
//                 message.reply(MESSAGES.INVALID_MAX_LEVEL(minLevel));
//                 return;
//               }
//             }
//           }

//           this.newItem[currentField] = value;
//           this.currentFieldIndex++;
//           this.requiredFieldsEntered = Math.min(this.currentFieldIndex, REQUIRED_FIELDS.length);

//           if (this.currentFieldIndex < REQUIRED_FIELDS.length) {
//             message.reply(MESSAGES.ADD_PROMPT(REQUIRED_FIELDS[this.currentFieldIndex]));
//           } else if (this.currentFieldIndex === REQUIRED_FIELDS.length) {
//             if (!this.validateRequiredFields()) {
//               message.reply(MESSAGES.MISSING_REQUIRED_FIELDS);
//               this.state = BotState.IDLE;
//               this.newItem = {};
//               this.currentFieldIndex = 0;
//               this.requiredFieldsEntered = 0;
//               return;
//             }
//             this.state = BotState.ADDING_CONFIRM;
//             message.reply(MESSAGES.ADD_CONFIRM);
//           } else {
//             const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
//             if (additionalFieldIndex < this.fieldsToAdd.length) {
//               message.reply(MESSAGES.ADD_PROMPT(this.fieldsToAdd[additionalFieldIndex]));
//             } else {
//               this.data.push(this.newItem);
//               await this.updateExcelFile();
//               message.reply(MESSAGES.SUCCESS_ADD);
//               this.newItem = {};
//               this.currentFieldIndex = 0;
//               this.requiredFieldsEntered = 0;
//               this.state = BotState.IDLE;
//             }
//           }
//           break;

//         case BotState.ADDING_CONFIRM:
//           console.log("[DEBUG] In ADDING_CONFIRM state.");
//           if (userMessage === "done") {
//             if (!this.validateRequiredFields()) {
//               message.reply(MESSAGES.MISSING_REQUIRED_FIELDS);
//               this.state = BotState.IDLE;
//               this.newItem = {};
//               this.currentFieldIndex = 0;
//               this.requiredFieldsEntered = 0;
//               return;
//             }
//             this.fillRemainingFields();
//             this.data.push(this.newItem);
//             await this.updateExcelFile();
//             message.reply(MESSAGES.SUCCESS_ADD);
//             this.newItem = {};
//             this.currentFieldIndex = 0;
//             this.requiredFieldsEntered = 0;
//             this.state = BotState.IDLE;
//           } else if (userMessage === "continue") {
//             this.state = BotState.ADDING_ITEM;
//             const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
//             if (additionalFieldIndex < this.fieldsToAdd.length) {
//               message.reply(MESSAGES.ADD_PROMPT(this.fieldsToAdd[additionalFieldIndex]));
//             } else {
//               this.data.push(this.newItem);
//               await this.updateExcelFile();
//               message.reply(MESSAGES.SUCCESS_ADD);
//               this.newItem = {};
//               this.currentFieldIndex = 0;
//               this.requiredFieldsEntered = 0;
//               this.state = BotState.IDLE;
//             }
//           } else {
//             message.reply(MESSAGES.INVALID_INPUT);
//           }
//           break;

//         case BotState.HANDLING_DUPLICATE:
//           console.log("[DEBUG] In HANDLING_DUPLICATE state.");
//           if (userMessage === "yes") {
//             // Show the existing item's details and transition to the UPDATING state
//             message.reply(MESSAGES.ITEM_DETAILS(this.currentItem));
//             this.state = BotState.UPDATING;
//             // Reset the newItem and related fields since we're not adding a new item
//             this.newItem = {};
//             this.currentFieldIndex = 0;
//             this.requiredFieldsEntered = 0;
//           } else if (userMessage === "no") {
//             message.reply("👍 Operation cancelled. Type 'hi' to start again.");
//             this.currentItem = null;
//             this.newItem = {};
//             this.currentFieldIndex = 0;
//             this.requiredFieldsEntered = 0;
//             this.state = BotState.IDLE;
//           } else {
//             message.reply(MESSAGES.INVALID_INPUT);
//           }
//           break;

//         default:
//           console.log("[DEBUG] Unknown state. Resetting to IDLE.");
//           this.state = BotState.IDLE;
//           message.reply(MESSAGES.ERROR);
//           break;
//       }
//     } catch (error) {
//       console.error("Error handling message:", error);
//       message.reply(MESSAGES.ERROR);
//       this.state = BotState.IDLE;
//     }
//   }
// }

// // Initialize the bot
// const bot = new InventoryBot();
// bot.listFiles();



















// const { Client, LocalAuth } = require("whatsapp-web.js");
// const qrcode = require("qrcode-terminal");
// const xlsx = require("xlsx");
// const fs = require("fs");
// const { google } = require("googleapis");
// require("dotenv").config();

// // Constants
// const FILE_PATH = "./a_data.xlsx";
// const SHEET_NAME_DEFAULT = "Sheet1";
// const SERIAL_NUMBER_FIELD = "S.No"; // The field name for the serial number in the Excel file
// const REQUIRED_FIELDS = ["Material Code", "Item Description"]; // Required fields for adding a new item
// const MESSAGES = {
//   WELCOME: "👋 Hi! Please choose an option:\n1. See/Update details of existing data\n2. Add a new item\n3. List all items\n4. Delete an item\nType 'help' for more commands.",
//   SEARCH_PROMPT: "🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to check details (or type 'cancel' to go back):",
//   DELETE_PROMPT: "🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to delete (or type 'cancel' to go back):",
//   ADD_PROMPT: (field) => `✏️ Enter ${field} (or type 'cancel' to go back):`,
//   ADD_CONFIRM: "✅ Material Code and Item Description added. Type 'done' to finish (remaining fields will be set to N/A) or 'continue' to add more details:",
//   UPDATE_PROMPT: (fields) => `🛠️ Which field would you like to update?\nAvailable fields:\n${fields}\n(or type 'cancel' to go back)`,
//   UPDATE_VALUE_PROMPT: (field) => `✏️ Enter the new value for ${field} (or type 'cancel' to go back):`,
//   SELECT_ITEM_PROMPT: (items) => `🔍 Multiple items found. Please select one:\n${items}\nReply with the number of the item (or type 'cancel' to go back).`,
//   ITEM_DETAILS: (item) => {
//     let reply = "📋 *Item Details:*\n";
//     for (const key in item) {
//       reply += `📌 *${key}:* ${item[key]}\n`;
//     }
//     return reply + "\n❓ Do you want to update any information? (yes/no)";
//   },
//   SUCCESS_ADD: "✅ New item added and Excel file updated on Google Drive!",
//   SUCCESS_DELETE: "✅ Item deleted successfully and Excel file updated on Google Drive!",
//   SUCCESS_UPDATE: (field, value) => `✅ ${field} updated successfully! New value: ${value}`,
//   UPDATE_CONFIRM: "⚙️ Do you want to update anything else? (yes/no)",
//   ERROR: "⚠️ An error occurred. Please try again or type 'hi' to restart.",
//   INVALID_INPUT: "❌ Invalid input. Please try again.",
//   NO_ITEMS_FOUND: "❌ No matching Material Code or Item Description found. Please try again.",
//   INVALID_SELECTION: "⚠️ Invalid selection. Please try again.",
//   CANCEL: "👍 Operation cancelled. Type 'hi' to start again.",
//   HELP: "📖 Available commands:\n- 'hi': Start the bot\n- 'help': Show this message\n- 'cancel': Cancel the current operation\nChoose an option:\n1. See/Update details\n2. Add a new item\n3. List all items\n4. Delete an item",
//   LOADING: "⏳ Loading data, please wait...",
//   LIST_ITEMS: (items) => `📋 *All Items:*\n${items}\nType 'hi' to go back to the main menu.`,
//   MISSING_REQUIRED_FIELDS: "❌ Both Material Code and Item Description are required. Item not added. Type 'hi' to start again.",
//   INVALID_MIN_LEVEL: (maxLevel) => `❌ Invalid input. Min Level must be less than or equal to Max Level (${maxLevel}). Please enter a valid value.`,
//   INVALID_MAX_LEVEL: (minLevel) => `❌ Invalid input. Max Level must be greater than or equal to Min Level (${minLevel}). Please enter a valid value.`,
//   INVALID_NUMERIC: "❌ Min Level and Max Level must be numeric values. Please enter a valid number.",
//   DUPLICATE_MATERIAL_CODE_PROMPT: (item) =>
//     `⚠️ The Material Code already exists:\n${item["S.No"]} - ${item["Material Code"]} - ${item["Item Description"]}\nWould you like to update this item? (yes/no)`,
// };

// // Google Drive Configuration
// const oauth2Client = new google.auth.OAuth2(
//   process.env.GOOGLE_CLIENT_ID,
//   process.env.GOOGLE_CLIENT_SECRET,
//   "urn:ietf:wg:oauth:2.0:oob"
// );
// oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
// const drive = google.drive({ version: "v3", auth: oauth2Client });

// // State Management
// const BotState = {
//   IDLE: "IDLE",
//   INITIAL_MENU: "INITIAL_MENU",
//   SEARCHING: "SEARCHING",
//   SELECTING_ITEM: "SELECTING_ITEM",
//   UPDATING: "UPDATING",
//   UPDATING_VALUE: "UPDATING_VALUE",
//   ADDING_ITEM: "ADDING_ITEM",
//   ADDING_CONFIRM: "ADDING_CONFIRM",
//   LISTING_ITEMS: "LISTING_ITEMS",
//   DELETING_ITEM: "DELETING_ITEM",
//   SELECTING_ITEM_TO_DELETE: "SELECTING_ITEM_TO_DELETE",
//   HANDLING_DUPLICATE: "HANDLING_DUPLICATE", // New state for handling duplicate Material Codes
// };

// // Class to manage bot state and data
// class InventoryBot {
//   constructor() {
//     this.state = BotState.IDLE;
//     this.data = [];
//     this.workbook = null;
//     this.sheetName = null;
//     this.worksheet = null;
//     this.headerRow = null;
//     this.currentItem = null;
//     this.foundItems = [];
//     this.newItem = {};
//     this.currentFieldIndex = 0;
//     this.fieldsToAdd = [];
//     this.requiredFieldsEntered = 0;
//     this.updateField = null;
//     this.client = new Client({
//       authStrategy: new LocalAuth(),
//       puppeteer: { args: ["--no-sandbox", "--disable-setuid-sandbox"] },
//     });

//     // Bind methods
//     this.handleMessage = this.handleMessage.bind(this);
//     this.downloadFileFromGoogleDrive = this.downloadFileFromGoogleDrive.bind(this);
//     this.uploadFileToGoogleDrive = this.uploadFileToGoogleDrive.bind(this);
//     this.loadExcelFile = this.loadExcelFile.bind(this);
//     this.updateExcelFile = this.updateExcelFile.bind(this);
//     this.listFiles = this.listFiles.bind(this);
//     this.generateSerialNumber = this.generateSerialNumber.bind(this);

//     // Initialize WhatsApp client
//     this.client.on("qr", (qr) => qrcode.generate(qr, { small: true }));
//     this.client.on("ready", () => console.log("✅ Client is ready!"));
//     this.client.on("message", this.handleMessage);
//     this.client.initialize();
//   }

//   // Generate the next serial number
//   generateSerialNumber() {
//     if (this.data.length === 0) {
//       return 1; // If the Excel file is empty, start with S.No 1
//     }
//     const serialNumbers = this.data
//       .map((item) => parseInt(item[SERIAL_NUMBER_FIELD], 10))
//       .filter((num) => !isNaN(num));
//     const maxSerialNumber = Math.max(...serialNumbers, 0);
//     return maxSerialNumber + 1;
//   }

//   // Google Drive: Download file with retry mechanism
//   async downloadFileFromGoogleDrive() {
//     const maxRetries = 3;
//     let attempt = 1;
//     while (attempt <= maxRetries) {
//       try {
//         console.log(`[DEBUG] Attempt ${attempt} to download file from Google Drive...`);
//         const response = await drive.files.get(
//           { fileId: process.env.GOOGLE_DRIVE_FILE_ID, alt: "media" },
//           { responseType: "stream" }
//         );
//         const writer = fs.createWriteStream(FILE_PATH);
//         response.data.pipe(writer);
//         return new Promise((resolve, reject) => {
//           writer.on("finish", () => {
//             console.log("File downloaded from Google Drive.");
//             resolve();
//           });
//           writer.on("error", (err) => {
//             console.error("Error downloading file:", err.message);
//             reject(err);
//           });
//         });
//       } catch (error) {
//         console.error(`Error downloading file from Google Drive (attempt ${attempt}):`, error.message);
//         if (attempt === maxRetries) throw error;
//         attempt++;
//         await new Promise(resolve => setTimeout(resolve, 2000)); // Wait 2 seconds before retrying
//       }
//     }
//   }

//   // Google Drive: Upload file
//   async uploadFileToGoogleDrive() {
//     try {
//       const fileContent = fs.createReadStream(FILE_PATH);
//       await drive.files.update({
//         fileId: process.env.GOOGLE_DRIVE_FILE_ID,
//         media: {
//           mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
//           body: fileContent,
//         },
//       });
//       console.log("File uploaded to Google Drive.");
//     } catch (error) {
//       console.error("Error uploading file to Google Drive:", error.message);
//       throw error;
//     }
//   }

//   // Google Drive: List files (for debugging)
//   async listFiles() {
//     try {
//       const response = await drive.files.list({
//         q: "name='a_data.xlsx'",
//         fields: "files(id, name, permissions, parents)",
//         spaces: "drive",
//         corpora: "user",
//         includeItemsFromAllDrives: true,
//         supportsAllDrives: true,
//       });
//       console.log("Files found:", response.data.files);
//       if (response.data.files.length === 0) {
//         console.log("No files found with name 'a_data.xlsx'. Checking case-insensitive...");
//         const responseCaseInsensitive = await drive.files.list({
//           q: "name contains 'a_data.xlsx'",
//           fields: "files(id, name, permissions, parents)",
//           spaces: "drive",
//           corpora: "user",
//           includeItemsFromAllDrives: true,
//           supportsAllDrives: true,
//         });
//         console.log("Files found (case-insensitive):", responseCaseInsensitive.data.files);
//       }
//     } catch (error) {
//       console.error("Error listing files:", error.response ? error.response.data : error.message);
//     }
//   }

//   // Excel: Load file
//   async loadExcelFile(message) {
//     try {
//       if (message) message.reply(MESSAGES.LOADING);
//       console.log("[DEBUG] Downloading file from Google Drive...");
//       // Force a fresh download by deleting the local file
//       if (fs.existsSync(FILE_PATH)) {
//         console.log("[DEBUG] Deleting existing local file to force fresh download...");
//         fs.unlinkSync(FILE_PATH);
//       }
//       await this.downloadFileFromGoogleDrive();

//       if (!fs.existsSync(FILE_PATH)) {
//         console.error("File not found locally. Creating an empty Excel file.");
//         this.workbook = xlsx.utils.book_new();
//         xlsx.utils.book_append_sheet(this.workbook, xlsx.utils.aoa_to_sheet([[]]), SHEET_NAME_DEFAULT);
//         xlsx.writeFile(this.workbook, FILE_PATH);
//       }

//       console.log("[DEBUG] Reading local Excel file...");
//       this.workbook = xlsx.readFile(FILE_PATH);
//       this.sheetName = this.workbook.SheetNames[0] || SHEET_NAME_DEFAULT;
//       console.log("[DEBUG] Using sheet:", this.sheetName);
//       this.worksheet = this.workbook.Sheets[this.sheetName];
//       this.headerRow = xlsx.utils.sheet_to_json(this.worksheet, { header: 1 })[0] || [];
//       console.log("[DEBUG] Header row:", this.headerRow);

//       this.data = xlsx.utils.sheet_to_json(this.worksheet);
//       this.data = this.data.map((entry) => {
//         let normalizedEntry = {};
//         this.headerRow.forEach((header) => {
//           const trimmedHeader = header?.toString().trim();
//           normalizedEntry[trimmedHeader] = entry[trimmedHeader] || "N/A";
//         });
//         return normalizedEntry;
//       });

//       console.log("📦 Loaded data from local file:", JSON.stringify(this.data, null, 2));
//     } catch (error) {
//       console.error("Error loading Excel file:", error.message);
//       throw error;
//     }
//   }

//   // Excel: Update file
//   async updateExcelFile() {
//     try {
//       const updatedWorksheet = xlsx.utils.json_to_sheet(this.data);
//       this.workbook.Sheets[this.sheetName] = updatedWorksheet;
//       xlsx.writeFile(this.workbook, FILE_PATH);
//       console.log("Excel file updated locally.");
//       await this.uploadFileToGoogleDrive();
//     } catch (error) {
//       console.error("Error updating Excel file:", error.message);
//       throw error;
//     }
//   }

//   // Fill remaining fields with "N/A"
//   fillRemainingFields() {
//     this.fieldsToAdd.forEach((field) => {
//       if (!(field in this.newItem)) {
//         this.newItem[field] = "N/A";
//       }
//     });
//   }

//   // Validate required fields
//   validateRequiredFields() {
//     return REQUIRED_FIELDS.every((field) => {
//       const value = this.newItem[field];
//       return value && value.trim() !== "" && value !== "N/A";
//     });
//   }

//   // Check if a value is defined (not "N/A" and not empty)
//   isDefined(value) {
//     return value && value !== "N/A" && value.toString().trim() !== "";
//   }

//   // WhatsApp: Handle messages
//   async handleMessage(message) {
//     try {
//       const userMessage = message.body.trim().toLowerCase();
//       console.log(`[DEBUG] Current state: ${this.state}, User message: ${userMessage}`);

//       // Handle "cancel" command
//       if (userMessage === "cancel" && this.state !== BotState.IDLE) {
//         console.log("[DEBUG] Cancel command received.");
//         // If required fields are entered, save the item before canceling
//         if (this.state === BotState.ADDING_ITEM || this.state === BotState.ADDING_CONFIRM) {
//           if (this.requiredFieldsEntered === REQUIRED_FIELDS.length && this.validateRequiredFields()) {
//             console.log("[DEBUG] Required fields entered, saving item before canceling.");
//             this.fillRemainingFields();
//             this.data.push(this.newItem);
//             await this.updateExcelFile();
//             message.reply(MESSAGES.SUCCESS_ADD);
//           } else {
//             console.log("[DEBUG] Required fields not fully entered, discarding item.");
//             message.reply(MESSAGES.CANCEL);
//           }
//         } else {
//           message.reply(MESSAGES.CANCEL);
//         }

//         // Reset state
//         this.state = BotState.IDLE;
//         this.currentItem = null;
//         this.foundItems = [];
//         this.newItem = {};
//         this.currentFieldIndex = 0;
//         this.requiredFieldsEntered = 0;
//         this.updateField = null;
//         return;
//       }

//       // Handle "help" command
//       if (userMessage === "help") {
//         console.log("[DEBUG] Help command received.");
//         message.reply(MESSAGES.HELP);
//         return;
//       }

//       // State machine
//       switch (this.state) {
//         case BotState.IDLE:
//           console.log("[DEBUG] In IDLE state.");
//           if (userMessage === "hi") {
//             await this.loadExcelFile(message);
//             if (this.data.length === 0) {
//               message.reply("⚠️ The Excel file is empty. Please add items using option 2.");
//               return;
//             }
//             // Exclude S.No and required fields from fieldsToAdd
//             this.fieldsToAdd = Object.keys(this.data[0] || {}).filter(
//               (field) =>
//                 field.toLowerCase() !== SERIAL_NUMBER_FIELD.toLowerCase() &&
//                 !REQUIRED_FIELDS.map(f => f.toLowerCase()).includes(field.toLowerCase())
//             );
//             message.reply(MESSAGES.WELCOME);
//             this.state = BotState.INITIAL_MENU;
//           }
//           break;

//         case BotState.INITIAL_MENU:
//           console.log("[DEBUG] In INITIAL_MENU state.");
//           if (userMessage === "1") {
//             this.state = BotState.SEARCHING;
//             message.reply(MESSAGES.SEARCH_PROMPT);
//           } else if (userMessage === "2") {
//             if (this.data.length === 0) {
//               message.reply("⚠️ The Excel file is empty. Please add at least one item first.");
//               return;
//             }
//             this.state = BotState.ADDING_ITEM;
//             this.currentFieldIndex = 0;
//             this.requiredFieldsEntered = 0;
//             this.newItem = {};
//             // Automatically generate S.No and add it to newItem
//             this.newItem[SERIAL_NUMBER_FIELD] = this.generateSerialNumber();
//             message.reply(MESSAGES.ADD_PROMPT(REQUIRED_FIELDS[this.currentFieldIndex]));
//           } else if (userMessage === "3") {
//             this.state = BotState.LISTING_ITEMS;
//             let itemsList = this.data.map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item["Material Code"]} - ${item["Item Description"]}`).join("\n");
//             message.reply(MESSAGES.LIST_ITEMS(itemsList));
//             this.state = BotState.IDLE;
//           } else if (userMessage === "4") {
//             this.state = BotState.DELETING_ITEM;
//             message.reply(MESSAGES.DELETE_PROMPT);
//           } else {
//             message.reply(MESSAGES.INVALID_INPUT);
//           }
//           break;

//         case BotState.SEARCHING:
//           console.log("[DEBUG] In SEARCHING state.");
//           const input = message.body.trim();
//           this.foundItems = this.data.filter(
//             (entry) =>
//               (entry["Material Code"] != null &&
//                 entry["Material Code"].toString().trim().toLowerCase().startsWith(input.toLowerCase())) ||
//               (typeof entry["Item Description"] === "string" &&
//                 entry["Item Description"].toLowerCase().includes(input.toLowerCase()))
//           );

//           if (this.foundItems.length === 1) {
//             this.currentItem = this.foundItems[0];
//             message.reply(MESSAGES.ITEM_DETAILS(this.currentItem));
//             this.state = BotState.UPDATING;
//           } else if (this.foundItems.length > 1) {
//             let itemsList = this.foundItems
//               .map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item["Material Code"]} - ${item["Item Description"]}`)
//               .join("\n");
//             message.reply(MESSAGES.SELECT_ITEM_PROMPT(itemsList));
//             this.state = BotState.SELECTING_ITEM;
//           } else {
//             message.reply(MESSAGES.NO_ITEMS_FOUND);
//           }
//           break;

//         case BotState.SELECTING_ITEM:
//           console.log("[DEBUG] In SELECTING_ITEM state.");
//           const selectedIndex = parseInt(userMessage) - 1;
//           if (!isNaN(selectedIndex) && selectedIndex >= 0 && selectedIndex < this.foundItems.length) {
//             this.currentItem = this.foundItems[selectedIndex];
//             message.reply(MESSAGES.ITEM_DETAILS(this.currentItem));
//             this.state = BotState.UPDATING;
//           } else {
//             message.reply(MESSAGES.INVALID_SELECTION);
//           }
//           break;

//         case BotState.DELETING_ITEM:
//           console.log("[DEBUG] In DELETING_ITEM state.");
//           const deleteInput = message.body.trim();
//           this.foundItems = this.data.filter(
//             (entry) =>
//               (entry["Material Code"] != null &&
//                 entry["Material Code"].toString().trim().toLowerCase().startsWith(deleteInput.toLowerCase())) ||
//               (typeof entry["Item Description"] === "string" &&
//                 entry["Item Description"].toLowerCase().includes(deleteInput.toLowerCase()))
//           );

//           if (this.foundItems.length === 1) {
//             this.currentItem = this.foundItems[0];
//             this.data = this.data.filter(
//               (entry) =>
//                 entry["Material Code"] !== this.currentItem["Material Code"] ||
//                 entry["Item Description"] !== this.currentItem["Item Description"]
//             );
//             await this.updateExcelFile();
//             message.reply(MESSAGES.SUCCESS_DELETE);
//             this.currentItem = null;
//             this.foundItems = [];
//             this.state = BotState.IDLE;
//           } else if (this.foundItems.length > 1) {
//             let itemsList = this.foundItems
//               .map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item["Material Code"]} - ${item["Item Description"]}`)
//               .join("\n");
//             message.reply(MESSAGES.SELECT_ITEM_PROMPT(itemsList));
//             this.state = BotState.SELECTING_ITEM_TO_DELETE;
//           } else {
//             message.reply(MESSAGES.NO_ITEMS_FOUND);
//             this.state = BotState.IDLE;
//           }
//           break;

//         case BotState.SELECTING_ITEM_TO_DELETE:
//           console.log("[DEBUG] In SELECTING_ITEM_TO_DELETE state.");
//           const deleteIndex = parseInt(userMessage) - 1;
//           if (!isNaN(deleteIndex) && deleteIndex >= 0 && deleteIndex < this.foundItems.length) {
//             this.currentItem = this.foundItems[deleteIndex];
//             this.data = this.data.filter(
//               (entry) =>
//                 entry["Material Code"] !== this.currentItem["Material Code"] ||
//                 entry["Item Description"] !== this.currentItem["Item Description"]
//             );
//             await this.updateExcelFile();
//             message.reply(MESSAGES.SUCCESS_DELETE);
//             this.currentItem = null;
//             this.foundItems = [];
//             this.state = BotState.IDLE;
//           } else {
//             message.reply(MESSAGES.INVALID_SELECTION);
//             this.state = BotState.IDLE;
//           }
//           break;

//         case BotState.UPDATING:
//           console.log("[DEBUG] In UPDATING state.");
//           if (userMessage === "yes") {
//             let availableFields = Object.keys(this.currentItem).map((field) => `📌 ${field}`).join("\n");
//             message.reply(MESSAGES.UPDATE_PROMPT(availableFields));
//           } else if (userMessage === "no") {
//             message.reply("👍 Thank you! Type 'hi' to start again.");
//             this.currentItem = null;
//             this.state = BotState.IDLE;
//           } else if (Object.keys(this.currentItem).map((key) => key.toLowerCase()).includes(userMessage)) {
//             this.updateField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === userMessage);
//             this.state = BotState.UPDATING_VALUE;
//             message.reply(MESSAGES.UPDATE_VALUE_PROMPT(this.updateField));
//           } else {
//             message.reply(MESSAGES.INVALID_INPUT);
//           }
//           break;

//         case BotState.UPDATING_VALUE:
//           console.log("[DEBUG] In UPDATING_VALUE state.");
//           const newValue = message.body.trim();
//           const fieldName = this.updateField;

//           // Validation for Min Level
//           if (fieldName.toLowerCase() === "min level") {
//             const maxLevelField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === "max level");
//             const maxLevel = maxLevelField ? this.currentItem[maxLevelField] : null;
//             if (this.isDefined(maxLevel)) {
//               if (isNaN(newValue) || isNaN(maxLevel)) {
//                 message.reply(MESSAGES.INVALID_NUMERIC);
//                 return;
//               }
//               if (Number(newValue) > Number(maxLevel)) {
//                 message.reply(MESSAGES.INVALID_MIN_LEVEL(maxLevel));
//                 return;
//               }
//             }
//           }
//           // Validation for Max Level
//           else if (fieldName.toLowerCase() === "max level") {
//             const minLevelField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === "min level");
//             const minLevel = minLevelField ? this.currentItem[minLevelField] : null;
//             if (this.isDefined(minLevel)) {
//               if (isNaN(newValue) || isNaN(minLevel)) {
//                 message.reply(MESSAGES.INVALID_NUMERIC);
//                 return;
//               }
//               if (Number(newValue) < Number(minLevel)) {
//                 message.reply(MESSAGES.INVALID_MAX_LEVEL(minLevel));
//                 return;
//               }
//             }
//           }

//           console.log(`[DEBUG] Updating ${fieldName} to ${newValue}`);
//           this.currentItem[this.updateField] = newValue;
//           await this.updateExcelFile();
//           message.reply(MESSAGES.SUCCESS_UPDATE(this.updateField, newValue));
//           message.reply(MESSAGES.UPDATE_CONFIRM);
//           this.state = BotState.UPDATING;
//           this.updateField = null;
//           console.log("[DEBUG] Transitioned back to UPDATING state.");
//           break;

//         case BotState.ADDING_ITEM:
//           console.log("[DEBUG] In ADDING_ITEM state.");
//           // Determine if we're collecting required fields or additional fields
//           let currentField;
//           if (this.currentFieldIndex < REQUIRED_FIELDS.length) {
//             currentField = REQUIRED_FIELDS[this.currentFieldIndex];
//           } else {
//             const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
//             currentField = this.fieldsToAdd[additionalFieldIndex];
//           }

//           const inputValue = message.body.trim();
//           let value = ["n/a", "na", "n/a", "n/a"].includes(inputValue.toLowerCase()) ? "" : inputValue;

//           // Check for duplicate Material Code when the user enters it
//           if (currentField === "Material Code") {
//             const existingItem = this.data.find(
//               (entry) =>
//                 entry["Material Code"] &&
//                 entry["Material Code"].toString().trim().toLowerCase() === value.toLowerCase()
//             );
//             if (existingItem) {
//               this.currentItem = existingItem; // Store the existing item
//               message.reply(MESSAGES.DUPLICATE_MATERIAL_CODE_PROMPT(existingItem));
//               this.state = BotState.HANDLING_DUPLICATE;
//               return;
//             }
//           }

//           // Validation for Min Level when adding a new item
//           if (currentField.toLowerCase() === "min level" && "Max Level" in this.newItem) {
//             const maxLevel = this.newItem["Max Level"];
//             if (this.isDefined(maxLevel)) {
//               if (isNaN(value) || isNaN(maxLevel)) {
//                 message.reply(MESSAGES.INVALID_NUMERIC);
//                 return;
//               }
//               if (Number(value) > Number(maxLevel)) {
//                 message.reply(MESSAGES.INVALID_MIN_LEVEL(maxLevel));
//                 return;
//               }
//             }
//           }
//           // Validation for Max Level when adding a new item
//           else if (currentField.toLowerCase() === "max level" && "Min Level" in this.newItem) {
//             const minLevel = this.newItem["Min Level"];
//             if (this.isDefined(minLevel)) {
//               if (isNaN(value) || isNaN(minLevel)) {
//                 message.reply(MESSAGES.INVALID_NUMERIC);
//                 return;
//               }
//               if (Number(value) < Number(minLevel)) {
//                 message.reply(MESSAGES.INVALID_MAX_LEVEL(minLevel));
//                 return;
//               }
//             }
//           }

//           this.newItem[currentField] = value;
//           this.currentFieldIndex++;
//           this.requiredFieldsEntered = Math.min(this.currentFieldIndex, REQUIRED_FIELDS.length);

//           if (this.currentFieldIndex < REQUIRED_FIELDS.length) {
//             message.reply(MESSAGES.ADD_PROMPT(REQUIRED_FIELDS[this.currentFieldIndex]));
//           } else if (this.currentFieldIndex === REQUIRED_FIELDS.length) {
//             if (!this.validateRequiredFields()) {
//               message.reply(MESSAGES.MISSING_REQUIRED_FIELDS);
//               this.state = BotState.IDLE;
//               this.newItem = {};
//               this.currentFieldIndex = 0;
//               this.requiredFieldsEntered = 0;
//               return;
//             }
//             this.state = BotState.ADDING_CONFIRM;
//             message.reply(MESSAGES.ADD_CONFIRM);
//           } else {
//             const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
//             if (additionalFieldIndex < this.fieldsToAdd.length) {
//               message.reply(MESSAGES.ADD_PROMPT(this.fieldsToAdd[additionalFieldIndex]));
//             } else {
//               this.data.push(this.newItem);
//               await this.updateExcelFile();
//               message.reply(MESSAGES.SUCCESS_ADD);
//               this.newItem = {};
//               this.currentFieldIndex = 0;
//               this.requiredFieldsEntered = 0;
//               this.state = BotState.IDLE;
//             }
//           }
//           break;

//         case BotState.ADDING_CONFIRM:
//           console.log("[DEBUG] In ADDING_CONFIRM state.");
//           if (userMessage === "done") {
//             if (!this.validateRequiredFields()) {
//               message.reply(MESSAGES.MISSING_REQUIRED_FIELDS);
//               this.state = BotState.IDLE;
//               this.newItem = {};
//               this.currentFieldIndex = 0;
//               this.requiredFieldsEntered = 0;
//               return;
//             }
//             this.fillRemainingFields();
//             this.data.push(this.newItem);
//             await this.updateExcelFile();
//             message.reply(MESSAGES.SUCCESS_ADD);
//             this.newItem = {};
//             this.currentFieldIndex = 0;
//             this.requiredFieldsEntered = 0;
//             this.state = BotState.IDLE;
//           } else if (userMessage === "continue") {
//             this.state = BotState.ADDING_ITEM;
//             const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
//             if (additionalFieldIndex < this.fieldsToAdd.length) {
//               message.reply(MESSAGES.ADD_PROMPT(this.fieldsToAdd[additionalFieldIndex]));
//             } else {
//               this.data.push(this.newItem);
//               await this.updateExcelFile();
//               message.reply(MESSAGES.SUCCESS_ADD);
//               this.newItem = {};
//               this.currentFieldIndex = 0;
//               this.requiredFieldsEntered = 0;
//               this.state = BotState.IDLE;
//             }
//           } else {
//             message.reply(MESSAGES.INVALID_INPUT);
//           }
//           break;

//         case BotState.HANDLING_DUPLICATE:
//           console.log("[DEBUG] In HANDLING_DUPLICATE state.");
//           if (userMessage === "yes") {
//             // Show the existing item's details and transition to the UPDATING state
//             message.reply(MESSAGES.ITEM_DETAILS(this.currentItem));
//             this.state = BotState.UPDATING;
//             // Reset the newItem and related fields since we're not adding a new item
//             this.newItem = {};
//             this.currentFieldIndex = 0;
//             this.requiredFieldsEntered = 0;
//           } else if (userMessage === "no") {
//             message.reply("👍 Operation cancelled. Type 'hi' to start again.");
//             this.currentItem = null;
//             this.newItem = {};
//             this.currentFieldIndex = 0;
//             this.requiredFieldsEntered = 0;
//             this.state = BotState.IDLE;
//           } else {
//             message.reply(MESSAGES.INVALID_INPUT);
//           }
//           break;

//         default:
//           console.log("[DEBUG] Unknown state. Resetting to IDLE.");
//           this.state = BotState.IDLE;
//           message.reply(MESSAGES.ERROR);
//           break;
//       }
//     } catch (error) {
//       console.error("Error handling message:", error);
//       message.reply(MESSAGES.ERROR);
//       this.state = BotState.IDLE;
//     }
//   }
// }

// // Initialize the bot
// const bot = new InventoryBot();
// bot.listFiles();








const { makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const xlsx = require('xlsx');
const fs = require('fs');
const { google } = require('googleapis');
const { Pool } = require('pg');
require('dotenv').config();

// Postgres client setup for Heroku
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }, // Required for Heroku Postgres
});

// Constants
const FILE_PATH = './inventory_register.xlsx';
const SHEET_NAME_DEFAULT = 'Sheet1';
const SERIAL_NUMBER_FIELD = 'S.No';
const REQUIRED_FIELDS = ['Material Code', 'Item Description'];

// Message Templates
const MESSAGES = {
  WELCOME: '👋 Hi! Please choose an option:\n1. See/Update details of existing data\n2. Add a new item\n3. List all items\n4. Delete an item\nType \'help\' for more commands.',
  SEARCH_PROMPT: '🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to check details (or type \'cancel\' to go back):',
  DELETE_PROMPT: '🔍 Please enter a 🔢 Material Code or 🏷️ Item Description to delete (or type \'cancel\' to go back):',
  ADD_PROMPT: (field) => `✏️ Enter ${field} (or type 'cancel' to go back):`,
  ADD_CONFIRM: '✅ Material Code and Item Description added. Type \'done\' to finish (remaining fields will be set to N/A) or \'continue\' to add more details:',
  UPDATE_PROMPT: (fields) => `🛠️ Which field would you like to update?\nAvailable fields:\n${fields}\n(or type 'cancel' to go back)`,
  UPDATE_VALUE_PROMPT: (field) => `✏️ Enter the new value for ${field} (or type 'cancel' to go back):`,
  SELECT_ITEM_PROMPT: (items) => `🔍 Multiple items found. Please select one:\n${items}\nReply with the number of the item (or type 'cancel' to go back).`,
  ITEM_DETAILS: (item) => {
    let reply = '📋 *Item Details:*\n';
    for (const key in item) {
      reply += `📌 *${key}:* ${item[key]}\n`;
    }
    return reply + '\n❓ Do you want to update any information? (yes/no)';
  },
  SUCCESS_ADD: '✅ New item added and Excel file updated on Google Drive!',
  SUCCESS_DELETE: '✅ Item deleted successfully and Excel file updated on Google Drive!',
  SUCCESS_UPDATE: (field, value) => `✅ ${field} updated successfully! New value: ${value}`,
  UPDATE_CONFIRM: '⚙️ Do you want to update anything else? (yes/no)',
  ERROR: '⚠️ An error occurred. Please try again or type \'hi\' to restart.',
  INVALID_INPUT: '❌ Invalid input. Please try again.',
  NO_ITEMS_FOUND: '❌ No matching Material Code or Item Description found. Please try again.',
  INVALID_SELECTION: '⚠️ Invalid selection. Please try again.',
  CANCEL: '👍 Operation cancelled. Type \'hi\' to start again.',
  HELP: '📖 Available commands:\n- \'hi\': Start the bot\n- \'help\': Show this message\n- \'cancel\': Cancel the current operation\nChoose an option:\n1. See/Update details\n2. Add a new item\n3. List all items\n4. Delete an item',
  LOADING: '⏳ Loading data, please wait...',
  LIST_ITEMS: (items) => `📋 *All Items:*\n${items}\nType 'hi' to go back to the main menu.`,
  MISSING_REQUIRED_FIELDS: '❌ Both Material Code and Item Description are required. Item not added. Type \'hi\' to start again.',
  INVALID_MIN_LEVEL: (maxLevel) => `❌ Invalid input. Min Level must be less than or equal to Max Level (${maxLevel}). Please enter a valid value.`,
  INVALID_MAX_LEVEL: (minLevel) => `❌ Invalid input. Max Level must be greater than or equal to Min Level (${minLevel}). Please enter a valid value.`,
  INVALID_NUMERIC: '❌ Min Level and Max Level must be numeric values. Please enter a valid number.',
  DUPLICATE_MATERIAL_CODE_PROMPT: (item) => `⚠️ The Material Code already exists:\n${item["S.No"]} - ${item["Material Code"]} - ${item["Item Description"]}\nWould you like to update this item? (yes/no)`,
};

// Google Drive Configuration
const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID,
  process.env.GOOGLE_CLIENT_SECRET,
  'urn:ietf:wg:oauth:2.0:oob'
);
oauth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });
const drive = google.drive({ version: 'v3', auth: oauth2Client });

// State Management
const BotState = {
  IDLE: 'IDLE',
  INITIAL_MENU: 'INITIAL_MENU',
  SEARCHING: 'SEARCHING',
  SELECTING_ITEM: 'SELECTING_ITEM',
  UPDATING: 'UPDATING',
  UPDATING_VALUE: 'UPDATING_VALUE',
  ADDING_ITEM: 'ADDING_ITEM',
  ADDING_CONFIRM: 'ADDING_CONFIRM',
  LISTING_ITEMS: 'LISTING_ITEMS',
  DELETING_ITEM: 'DELETING_ITEM',
  SELECTING_ITEM_TO_DELETE: 'SELECTING_ITEM_TO_DELETE',
  HANDLING_DUPLICATE: 'HANDLING_DUPLICATE',
};

// Custom auth state handler for Postgres
// Replace the existing usePostgresAuthState function with this corrected version
async function usePostgresAuthState() {
  const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: { rejectUnauthorized: false }
  });

  // Ensure the table exists
  await pool.query(`
    CREATE TABLE IF NOT EXISTS auth_state (
      key TEXT PRIMARY KEY,
      value JSONB
    )
  `);

  const loadState = async () => {
    try {
      const result = await pool.query('SELECT value FROM auth_state WHERE key = $1', ['creds']);
      if (result.rows[0]?.value) {
        const creds = JSON.parse(result.rows[0].value);
        // Verify that creds.me exists
        if (creds && creds.me && creds.me.id) {
          return { creds };
        }
      }
      // Return empty creds to trigger QR generation
      return { creds: null };
    } catch (error) {
      console.error('Error loading auth state from Postgres:', error);
      return { creds: null };
    }
  };

  const saveCreds = async (creds) => {
    try {
      await pool.query(
        'INSERT INTO auth_state (key, value) VALUES ($1, $2) ON CONFLICT (key) DO UPDATE SET value = $2',
        ['creds', JSON.stringify(creds)]
      );
    } catch (error) {
      console.error('Error saving credentials to Postgres:', error);
      throw error;
    }
  };

  // Load initial state
  const state = await loadState();

  return {
    state,
    saveCreds
  };
}
// Class to manage bot state and data
class InventoryBot {
  constructor() {
    this.state = BotState.IDLE;
    this.data = [];
    this.workbook = null;
    this.sheetName = null;
    this.worksheet = null;
    this.headerRow = null;
    this.currentItem = null;
    this.foundItems = [];
    this.newItem = {};
    this.currentFieldIndex = 0;
    this.fieldsToAdd = [];
    this.requiredFieldsEntered = 0;
    this.updateField = null;
    this.sock = null;

    // Bind methods
    this.start = this.start.bind(this);
    this.handleMessage = this.handleMessage.bind(this);
    this.downloadFileFromGoogleDrive = this.downloadFileFromGoogleDrive.bind(this);
    this.uploadFileToGoogleDrive = this.uploadFileToGoogleDrive.bind(this);
    this.loadExcelFile = this.loadExcelFile.bind(this);
    this.updateExcelFile = this.updateExcelFile.bind(this);
    this.generateSerialNumber = this.generateSerialNumber.bind(this);
    this.sendMessage = this.sendMessage.bind(this);

    // Start the bot
    this.start();
  }

   // Update the start() function's auth handling section
async start() {
  let auth;
  if (process.env.DATABASE_URL) {
    // Use Postgres auth
    const { state, saveCreds } = await usePostgresAuthState();
    auth = {
      state: state,
      saveCreds: saveCreds
    };
  } else {
    // Use file-based auth locally
    const { state, saveCreds } = await useMultiFileAuthState('auth_info');
    auth = {
      state: state,
      saveCreds: saveCreds
    };
  }

  this.sock = makeWASocket({
    auth: auth.state,
    printQRInTerminal: true,
  });

  this.sock.ev.on('creds.update', auth.saveCreds);

    this.sock.ev.on('connection.update', (update) => {
      const { connection, lastDisconnect, qr } = update;
      if (qr) {
        console.log('QR Code for authentication:', qr);
      }
      if (connection === 'close') {
        const shouldReconnect = lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut;
        console.log('Connection closed due to ', lastDisconnect?.error, ', reconnecting ', shouldReconnect);
        if (shouldReconnect) {
          this.start();
        }
      } else if (connection === 'open') {
        console.log('✅ Client is ready!');
      }
    });

    this.sock.ev.on('messages.upsert', async ({ messages }) => {
      const msg = messages[0];
      if (!msg.message || msg.key.fromMe) return; // Ignore messages sent by the bot itself
      await this.handleMessage(msg);
    });
  }

  // Helper method to send messages
  async sendMessage(to, msg) {
    try {
      console.log('[DEBUG] Sending text message:', msg);
      await this.sock.sendMessage(to, { text: msg });
    } catch (error) {
      console.error('[ERROR] Failed to send message:', error);
      await this.sock.sendMessage(to, { text: MESSAGES.ERROR });
    }
  }

  // Generate the next serial number
  generateSerialNumber() {
    if (this.data.length === 0) {
      return 1;
    }
    const serialNumbers = this.data
      .map((item) => parseInt(item[SERIAL_NUMBER_FIELD], 10))
      .filter((num) => !isNaN(num));
    const maxSerialNumber = Math.max(...serialNumbers, 0);
    return maxSerialNumber + 1;
  }

  // Google Drive: Download file with retry mechanism
  async downloadFileFromGoogleDrive() {
    const maxRetries = 3;
    let attempt = 1;
    while (attempt <= maxRetries) {
      try {
        console.log(`[DEBUG] Attempt ${attempt} to download file from Google Drive...`);
        const response = await drive.files.get(
          { fileId: process.env.GOOGLE_DRIVE_FILE_ID, alt: 'media' },
          { responseType: 'stream' }
        );
        const writer = fs.createWriteStream(FILE_PATH);
        response.data.pipe(writer);
        return new Promise((resolve, reject) => {
          writer.on('finish', () => {
            console.log('File downloaded from Google Drive.');
            resolve();
          });
          writer.on('error', (err) => {
            console.error('Error downloading file:', err.message);
            reject(err);
          });
        });
      } catch (error) {
        console.error(`Error downloading file from Google Drive (attempt ${attempt}):`, error.message);
        if (attempt === maxRetries) throw error;
        attempt++;
        await new Promise(resolve => setTimeout(resolve, 2000));
      }
    }
  }

  // Google Drive: Upload file
  async uploadFileToGoogleDrive() {
    try {
      const fileContent = fs.createReadStream(FILE_PATH);
      await drive.files.update({
        fileId: process.env.GOOGLE_DRIVE_FILE_ID,
        media: {
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          body: fileContent,
        },
      });
      console.log('File uploaded to Google Drive.');
    } catch (error) {
      console.error('Error uploading file to Google Drive:', error.message);
      throw error;
    }
  }

  // Excel: Load file
  async loadExcelFile(to) {
    try {
      if (to) await this.sendMessage(to, MESSAGES.LOADING);
      console.log('[DEBUG] Downloading file from Google Drive...');
      if (fs.existsSync(FILE_PATH)) {
        console.log('[DEBUG] Deleting existing local file to force fresh download...');
        fs.unlinkSync(FILE_PATH);
      }
      await this.downloadFileFromGoogleDrive();

      if (!fs.existsSync(FILE_PATH)) {
        console.error('File not found locally. Creating an empty Excel file.');
        this.workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(this.workbook, xlsx.utils.aoa_to_sheet([[]]), SHEET_NAME_DEFAULT);
        xlsx.writeFile(this.workbook, FILE_PATH);
      }

      console.log('[DEBUG] Reading local Excel file...');
      this.workbook = xlsx.readFile(FILE_PATH);
      this.sheetName = this.workbook.SheetNames[0] || SHEET_NAME_DEFAULT;
      console.log('[DEBUG] Using sheet:', this.sheetName);
      this.worksheet = this.workbook.Sheets[this.sheetName];
      this.headerRow = xlsx.utils.sheet_to_json(this.worksheet, { header: 1 })[0] || [];
      console.log('[DEBUG] Header row:', this.headerRow);

      this.data = xlsx.utils.sheet_to_json(this.worksheet);
      this.data = this.data.map((entry) => {
        let normalizedEntry = {};
        this.headerRow.forEach((header) => {
          const trimmedHeader = header?.toString().trim();
          normalizedEntry[trimmedHeader] = entry[trimmedHeader] || 'N/A';
        });
        return normalizedEntry;
      });

      console.log('📦 Loaded data from local file:', JSON.stringify(this.data, null, 2));
    } catch (error) {
      console.error('Error loading Excel file:', error.message);
      throw error;
    }
  }

  // Excel: Update file
  async updateExcelFile() {
    try {
      const updatedWorksheet = xlsx.utils.json_to_sheet(this.data);
      this.workbook.Sheets[this.sheetName] = updatedWorksheet;
      xlsx.writeFile(this.workbook, FILE_PATH);
      console.log('Excel file updated locally.');
      await this.uploadFileToGoogleDrive();
    } catch (error) {
      console.error('Error updating Excel file:', error.message);
      throw error;
    }
  }

  // Fill remaining fields with "N/A"
  fillRemainingFields() {
    this.fieldsToAdd.forEach((field) => {
      if (!(field in this.newItem)) {
        this.newItem[field] = 'N/A';
      }
    });
  }

  // Validate required fields
  validateRequiredFields() {
    return REQUIRED_FIELDS.every((field) => {
      const value = this.newItem[field];
      return value && value.trim() !== '' && value !== 'N/A';
    });
  }

  // Check if a value is defined (not "N/A" and not empty)
  isDefined(value) {
    return value && value !== 'N/A' && value.toString().trim() !== '';
  }

  // WhatsApp: Handle messages
  async handleMessage(msg) {
    try {
      const to = msg.key.remoteJid;
      let userMessage = msg.message?.conversation || msg.message?.extendedTextMessage?.text || '';
      userMessage = userMessage.trim().toLowerCase();

      console.log(`[DEBUG] Current state: ${this.state}, User message: ${userMessage}`);

      // Handle "cancel" command
      if (userMessage === 'cancel' && this.state !== BotState.IDLE) {
        console.log('[DEBUG] Cancel command received.');
        if (this.state === BotState.ADDING_ITEM || this.state === BotState.ADDING_CONFIRM) {
          if (this.requiredFieldsEntered === REQUIRED_FIELDS.length && this.validateRequiredFields()) {
            console.log('[DEBUG] Required fields entered, saving item before canceling.');
            this.fillRemainingFields();
            this.data.push(this.newItem);
            await this.updateExcelFile();
            await this.sendMessage(to, MESSAGES.SUCCESS_ADD);
          } else {
            console.log('[DEBUG] Required fields not fully entered, discarding item.');
            await this.sendMessage(to, MESSAGES.CANCEL);
          }
        } else {
          await this.sendMessage(to, MESSAGES.CANCEL);
        }

        // Reset state
        this.state = BotState.IDLE;
        this.currentItem = null;
        this.foundItems = [];
        this.newItem = {};
        this.currentFieldIndex = 0;
        this.requiredFieldsEntered = 0;
        this.updateField = null;
        return;
      }

      // Handle "help" command
      if (userMessage === 'help') {
        console.log('[DEBUG] Help command received.');
        await this.sendMessage(to, MESSAGES.HELP);
        return;
      }

      // State machine
      switch (this.state) {
        case BotState.IDLE:
          console.log('[DEBUG] In IDLE state.');
          if (userMessage === 'hi') {
            await this.loadExcelFile(to);
            if (this.data.length === 0) {
              await this.sendMessage(to, '⚠️ The Excel file is empty. Please add items using option 2.');
              return;
            }
            this.fieldsToAdd = Object.keys(this.data[0] || {}).filter(
              (field) =>
                field.toLowerCase() !== SERIAL_NUMBER_FIELD.toLowerCase() &&
                !REQUIRED_FIELDS.map(f => f.toLowerCase()).includes(field.toLowerCase())
            );
            await this.sendMessage(to, MESSAGES.WELCOME);
            this.state = BotState.INITIAL_MENU;
          }
          break;

        case BotState.INITIAL_MENU:
          console.log('[DEBUG] In INITIAL_MENU state.');
          if (userMessage === '1') {
            this.state = BotState.SEARCHING;
            await this.sendMessage(to, MESSAGES.SEARCH_PROMPT);
          } else if (userMessage === '2') {
            if (this.data.length === 0) {
              await this.sendMessage(to, '⚠️ The Excel file is empty. Please add at least one item first.');
              return;
            }
            this.state = BotState.ADDING_ITEM;
            this.currentFieldIndex = 0;
            this.requiredFieldsEntered = 0;
            this.newItem = {};
            this.newItem[SERIAL_NUMBER_FIELD] = this.generateSerialNumber();
            await this.sendMessage(to, MESSAGES.ADD_PROMPT(REQUIRED_FIELDS[this.currentFieldIndex]));
          } else if (userMessage === '3') {
            this.state = BotState.LISTING_ITEMS;
            let itemsList = this.data.map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item['Material Code']} - ${item['Item Description']}`).join('\n');
            await this.sendMessage(to, MESSAGES.LIST_ITEMS(itemsList));
            this.state = BotState.IDLE;
          } else if (userMessage === '4') {
            this.state = BotState.DELETING_ITEM;
            await this.sendMessage(to, MESSAGES.DELETE_PROMPT);
          } else {
            await this.sendMessage(to, MESSAGES.INVALID_INPUT);
          }
          break;

        case BotState.SEARCHING:
          console.log('[DEBUG] In SEARCHING state.');
          const input = msg.message?.conversation || msg.message?.extendedTextMessage?.text;
          this.foundItems = this.data.filter(
            (entry) =>
              (entry['Material Code'] != null &&
                entry['Material Code'].toString().trim().toLowerCase().startsWith(input.toLowerCase())) ||
              (typeof entry['Item Description'] === 'string' &&
                entry['Item Description'].toLowerCase().includes(input.toLowerCase()))
          );

          if (this.foundItems.length === 1) {
            this.currentItem = this.foundItems[0];
            await this.sendMessage(to, MESSAGES.ITEM_DETAILS(this.currentItem));
            this.state = BotState.UPDATING;
          } else if (this.foundItems.length > 1) {
            let itemsList = this.foundItems
              .map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item['Material Code']} - ${item['Item Description']}`)
              .join('\n');
            await this.sendMessage(to, MESSAGES.SELECT_ITEM_PROMPT(itemsList));
            this.state = BotState.SELECTING_ITEM;
          } else {
            await this.sendMessage(to, MESSAGES.NO_ITEMS_FOUND);
          }
          break;

        case BotState.SELECTING_ITEM:
          console.log('[DEBUG] In SELECTING_ITEM state.');
          const selectedIndex = parseInt(userMessage) - 1;
          if (!isNaN(selectedIndex) && selectedIndex >= 0 && selectedIndex < this.foundItems.length) {
            this.currentItem = this.foundItems[selectedIndex];
            await this.sendMessage(to, MESSAGES.ITEM_DETAILS(this.currentItem));
            this.state = BotState.UPDATING;
          } else {
            await this.sendMessage(to, MESSAGES.INVALID_SELECTION);
          }
          break;

        case BotState.DELETING_ITEM:
          console.log('[DEBUG] In DELETING_ITEM state.');
          const deleteInput = msg.message?.conversation || msg.message?.extendedTextMessage?.text;
          this.foundItems = this.data.filter(
            (entry) =>
              (entry['Material Code'] != null &&
                entry['Material Code'].toString().trim().toLowerCase().startsWith(deleteInput.toLowerCase())) ||
              (typeof entry['Item Description'] === 'string' &&
                entry['Item Description'].toLowerCase().includes(deleteInput.toLowerCase()))
          );

          if (this.foundItems.length === 1) {
            this.currentItem = this.foundItems[0];
            this.data = this.data.filter(
              (entry) =>
                entry['Material Code'] !== this.currentItem['Material Code'] ||
                entry['Item Description'] !== this.currentItem['Item Description']
            );
            await this.updateExcelFile();
            await this.sendMessage(to, MESSAGES.SUCCESS_DELETE);
            this.currentItem = null;
            this.foundItems = [];
            this.state = BotState.IDLE;
          } else if (this.foundItems.length > 1) {
            let itemsList = this.foundItems
              .map((item, index) => `${index + 1}. ${item[SERIAL_NUMBER_FIELD]} - ${item['Material Code']} - ${item['Item Description']}`)
              .join('\n');
            await this.sendMessage(to, MESSAGES.SELECT_ITEM_PROMPT(itemsList));
            this.state = BotState.SELECTING_ITEM_TO_DELETE;
          } else {
            await this.sendMessage(to, MESSAGES.NO_ITEMS_FOUND);
            this.state = BotState.IDLE;
          }
          break;

        case BotState.SELECTING_ITEM_TO_DELETE:
          console.log('[DEBUG] In SELECTING_ITEM_TO_DELETE state.');
          const deleteIndex = parseInt(userMessage) - 1;
          if (!isNaN(deleteIndex) && deleteIndex >= 0 && deleteIndex < this.foundItems.length) {
            this.currentItem = this.foundItems[deleteIndex];
            this.data = this.data.filter(
              (entry) =>
                entry['Material Code'] !== this.currentItem['Material Code'] ||
                entry['Item Description'] !== this.currentItem['Item Description']
            );
            await this.updateExcelFile();
            await this.sendMessage(to, MESSAGES.SUCCESS_DELETE);
            this.currentItem = null;
            this.foundItems = [];
            this.state = BotState.IDLE;
          } else {
            await this.sendMessage(to, MESSAGES.INVALID_SELECTION);
            this.state = BotState.IDLE;
          }
          break;

        case BotState.UPDATING:
          console.log('[DEBUG] In UPDATING state.');
          if (userMessage === 'yes') {
            let availableFields = Object.keys(this.currentItem).map((field) => `📌 ${field}`).join('\n');
            await this.sendMessage(to, MESSAGES.UPDATE_PROMPT(availableFields));
          } else if (userMessage === 'no') {
            await this.sendMessage(to, '👍 Thank you! Type \'hi\' to start again.');
            this.currentItem = null;
            this.state = BotState.IDLE;
          } else if (Object.keys(this.currentItem).map((key) => key.toLowerCase()).includes(userMessage)) {
            this.updateField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === userMessage);
            this.state = BotState.UPDATING_VALUE;
            await this.sendMessage(to, MESSAGES.UPDATE_VALUE_PROMPT(this.updateField));
          } else {
            await this.sendMessage(to, MESSAGES.INVALID_INPUT);
          }
          break;

        case BotState.UPDATING_VALUE:
          console.log('[DEBUG] In UPDATING_VALUE state.');
          const newValue = msg.message?.conversation || msg.message?.extendedTextMessage?.text;
          const fieldName = this.updateField;

          if (fieldName.toLowerCase() === 'min level') {
            const maxLevelField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === 'max level');
            const maxLevel = maxLevelField ? this.currentItem[maxLevelField] : null;
            if (this.isDefined(maxLevel)) {
              if (isNaN(newValue) || isNaN(maxLevel)) {
                await this.sendMessage(to, MESSAGES.INVALID_NUMERIC);
                return;
              }
              if (Number(newValue) > Number(maxLevel)) {
                await this.sendMessage(to, MESSAGES.INVALID_MIN_LEVEL(maxLevel));
                return;
              }
            }
          } else if (fieldName.toLowerCase() === 'max level') {
            const minLevelField = Object.keys(this.currentItem).find((key) => key.toLowerCase() === 'min level');
            const minLevel = minLevelField ? this.currentItem[minLevelField] : null;
            if (this.isDefined(minLevel)) {
              if (isNaN(newValue) || isNaN(minLevel)) {
                await this.sendMessage(to, MESSAGES.INVALID_NUMERIC);
                return;
              }
              if (Number(newValue) < Number(minLevel)) {
                await this.sendMessage(to, MESSAGES.INVALID_MAX_LEVEL(minLevel));
                return;
              }
            }
          }

          console.log(`[DEBUG] Updating ${fieldName} to ${newValue}`);
          this.currentItem[this.updateField] = newValue;
          await this.updateExcelFile();
          await this.sendMessage(to, MESSAGES.SUCCESS_UPDATE(this.updateField, newValue));
          await this.sendMessage(to, MESSAGES.UPDATE_CONFIRM);
          this.state = BotState.UPDATING;
          this.updateField = null;
          console.log('[DEBUG] Transitioned back to UPDATING state.');
          break;

        case BotState.ADDING_ITEM:
          console.log('[DEBUG] In ADDING_ITEM state.');
          let currentField;
          if (this.currentFieldIndex < REQUIRED_FIELDS.length) {
            currentField = REQUIRED_FIELDS[this.currentFieldIndex];
          } else {
            const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
            currentField = this.fieldsToAdd[additionalFieldIndex];
          }

          const inputValue = msg.message?.conversation || msg.message?.extendedTextMessage?.text;
          let value = ['n/a', 'na', 'n/a', 'n/a'].includes(inputValue.toLowerCase()) ? '' : inputValue;

          if (currentField === 'Material Code') {
            const existingItem = this.data.find(
              (entry) =>
                entry['Material Code'] &&
                entry['Material Code'].toString().trim().toLowerCase() === value.toLowerCase()
            );
            if (existingItem) {
              this.currentItem = existingItem;
              await this.sendMessage(to, MESSAGES.DUPLICATE_MATERIAL_CODE_PROMPT(existingItem));
              this.state = BotState.HANDLING_DUPLICATE;
              return;
            }
          }

          if (currentField.toLowerCase() === 'min level' && 'Max Level' in this.newItem) {
            const maxLevel = this.newItem['Max Level'];
            if (this.isDefined(maxLevel)) {
              if (isNaN(value) || isNaN(maxLevel)) {
                await this.sendMessage(to, MESSAGES.INVALID_NUMERIC);
                return;
              }
              if (Number(value) > Number(maxLevel)) {
                await this.sendMessage(to, MESSAGES.INVALID_MIN_LEVEL(maxLevel));
                return;
              }
            }
          } else if (currentField.toLowerCase() === 'max level' && 'Min Level' in this.newItem) {
            const minLevel = this.newItem['Min Level'];
            if (this.isDefined(minLevel)) {
              if (isNaN(value) || isNaN(minLevel)) {
                await this.sendMessage(to, MESSAGES.INVALID_NUMERIC);
                return;
              }
              if (Number(value) < Number(minLevel)) {
                await this.sendMessage(to, MESSAGES.INVALID_MAX_LEVEL(minLevel));
                return;
              }
            }
          }

          this.newItem[currentField] = value;
          this.currentFieldIndex++;
          this.requiredFieldsEntered = Math.min(this.currentFieldIndex, REQUIRED_FIELDS.length);

          if (this.currentFieldIndex < REQUIRED_FIELDS.length) {
            await this.sendMessage(to, MESSAGES.ADD_PROMPT(REQUIRED_FIELDS[this.currentFieldIndex]));
          } else if (this.currentFieldIndex === REQUIRED_FIELDS.length) {
            if (!this.validateRequiredFields()) {
              await this.sendMessage(to, MESSAGES.MISSING_REQUIRED_FIELDS);
              this.state = BotState.IDLE;
              this.newItem = {};
              this.currentFieldIndex = 0;
              this.requiredFieldsEntered = 0;
              return;
            }
            this.state = BotState.ADDING_CONFIRM;
            await this.sendMessage(to, MESSAGES.ADD_CONFIRM);
          } else {
            const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
            if (additionalFieldIndex < this.fieldsToAdd.length) {
              await this.sendMessage(to, MESSAGES.ADD_PROMPT(this.fieldsToAdd[additionalFieldIndex]));
            } else {
              this.data.push(this.newItem);
              await this.updateExcelFile();
              await this.sendMessage(to, MESSAGES.SUCCESS_ADD);
              this.newItem = {};
              this.currentFieldIndex = 0;
              this.requiredFieldsEntered = 0;
              this.state = BotState.IDLE;
            }
          }
          break;

        case BotState.ADDING_CONFIRM:
          console.log('[DEBUG] In ADDING_CONFIRM state.');
          if (userMessage === 'done') {
            if (!this.validateRequiredFields()) {
              await this.sendMessage(to, MESSAGES.MISSING_REQUIRED_FIELDS);
              this.state = BotState.IDLE;
              this.newItem = {};
              this.currentFieldIndex = 0;
              this.requiredFieldsEntered = 0;
              return;
            }
            this.fillRemainingFields();
            this.data.push(this.newItem);
            await this.updateExcelFile();
            await this.sendMessage(to, MESSAGES.SUCCESS_ADD);
            this.newItem = {};
            this.currentFieldIndex = 0;
            this.requiredFieldsEntered = 0;
            this.state = BotState.IDLE;
          } else if (userMessage === 'continue') {
            this.state = BotState.ADDING_ITEM;
            const additionalFieldIndex = this.currentFieldIndex - REQUIRED_FIELDS.length;
            if (additionalFieldIndex < this.fieldsToAdd.length) {
              await this.sendMessage(to, MESSAGES.ADD_PROMPT(this.fieldsToAdd[additionalFieldIndex]));
            } else {
              this.data.push(this.newItem);
              await this.updateExcelFile();
              await this.sendMessage(to, MESSAGES.SUCCESS_ADD);
              this.newItem = {};
              this.currentFieldIndex = 0;
              this.requiredFieldsEntered = 0;
              this.state = BotState.IDLE;
            }
          } else {
            await this.sendMessage(to, MESSAGES.INVALID_INPUT);
          }
          break;

        case BotState.HANDLING_DUPLICATE:
          console.log('[DEBUG] In HANDLING_DUPLICATE state.');
          if (userMessage === 'yes') {
            await this.sendMessage(to, MESSAGES.ITEM_DETAILS(this.currentItem));
            this.state = BotState.UPDATING;
            this.newItem = {};
            this.currentFieldIndex = 0;
            this.requiredFieldsEntered = 0;
          } else if (userMessage === 'no') {
            await this.sendMessage(to, '👍 Operation cancelled. Type \'hi\' to start again.');
            this.currentItem = null;
            this.newItem = {};
            this.currentFieldIndex = 0;
            this.requiredFieldsEntered = 0;
            this.state = BotState.IDLE;
          } else {
            await this.sendMessage(to, MESSAGES.INVALID_INPUT);
          }
          break;

        default:
          console.log('[DEBUG] Unknown state. Resetting to IDLE.');
          this.state = BotState.IDLE;
          await this.sendMessage(to, MESSAGES.ERROR);
          break;
      }
    } catch (error) {
      console.error('Error handling message:', error);
      await this.sendMessage(msg.key.remoteJid, MESSAGES.ERROR);
      this.state = BotState.IDLE;
    }
  }
}

// Start the bot
new InventoryBot();

// Keep the dyno alive (Heroku requires a web process)
const express = require('express');
const app = express();
app.get('/', (req, res) => res.send('Bot is running!'));
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));