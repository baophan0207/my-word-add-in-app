# Word Add-in Application

This project consists of three main components: a React application, a Node.js server, and a Word add-in.

## Project Structure

my-word-add-in-app/
├── react-app/ # React frontend application
├── node-server/ # Node.js backend server
└── word-add-in/ # Word add-in files

## Getting Started

### React Application (react-app)

1. Navigate to the React app directory:
   bash
   cd react-app

2. Install dependencies:

```bash
npm install
```

3. Start the development server:

```bash
npm start
```

The React app will run on http://localhost:3000

### Node.js Server (node-server)

1. Navigate to the Node.js server directory:

```bash
cd node-server
```

2. Install dependencies:

```bash
npm install
```

3. Start the server:

```bash
npm start
```

The server will run on http://localhost:3001

### Word Add-in (word-add-in)

The Word add-in can be installed in two ways:

#### Option 1: Using the React Application

1. Ensure both the React app and Node.js server are running
2. Navigate to the React application in your browser
3. Follow the in-app instructions to install the add-in

#### Option 2: Manual Installation using PowerShell

1. Navigate to the word-add-in directory:

```bash
cd word-add-in
```

2. Run PowerShell as Administrator

3. Set the execution policy to run the script:

```bash
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

4. Run the installation script:

```bash
.\install-addin.ps1
```

5. Follow any on-screen prompts

### Verifying the Installation

After installation, you can verify the add-in is properly installed:

1. Open Microsoft Word
2. Check the Home tab for your add-in's button
3. If you don't see the button:
   - Click File > Options > Add-ins
   - Look for your add-in in the list

## Troubleshooting

### Word Add-in Not Appearing

If the add-in doesn't appear in Word after installation:

1. Close all Word instances
2. Clear the Office cache:
   - Delete contents of: %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
   - Delete contents of: %USERPROFILE%\AppData\Local\Microsoft\Office\16.0\WefCache
3. Restart Word

### Installation Errors

If you encounter errors during installation:

1. Ensure you're running PowerShell as Administrator
2. Verify the manifest.xml file exists in the correct location
3. Check that all paths in the manifest file are correct

### Removing the Add-in Manually

If you need to remove the add-in manually:

1. Open Registry Editor:

   - Press `Win + R`
   - Type `regedit` and press Enter
   - Click "Yes" if prompted by UAC

2. Navigate to:

   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins
   ```

3. Find your add-in:

   - Look for the ID: `f85491a7-0cf8-4950-b18c-d85ae9970d61`
   - Right-click on it
   - Select "Delete"
   - Click "Yes" to confirm

4. Close Registry Editor

5. Restart Word for the changes to take effect

## Available Scripts

### React App

- `npm start`: Runs the app in development mode
- `npm test`: Launches the test runner
- `npm run build`: Builds the app for production

### Node Server

- `node server.js`: Starts the Node.js server

## Learn More

- [Create React App documentation](https://facebook.github.io/create-react-app/docs/getting-started)
- [Office Add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
