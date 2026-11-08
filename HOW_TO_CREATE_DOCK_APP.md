# How to Create a macOS Dock App for VCard Merger

**Make your Contact Merger app launchable from the Dock - no Terminal needed!**

---

## Method 1: Automator Application ⭐ RECOMMENDED

**What you need:** Nothing! Automator is built into macOS.

**Time:** 3-5 minutes

**Result:** A real macOS app you can launch from the Dock with a custom icon.

---

### Step-by-Step Instructions

#### Part A: Create the Application

**1. Open Automator**
- Press `Cmd + Space` to open Spotlight
- Type "Automator" and press Enter
- Or find it in `/Applications/Automator.app`

**2. Create New Application**
- Click "New Document" (if Automator asks)
- Select **"Application"** from the list of document types
- Click "Choose"

**3. Add Shell Script Action**
- In the left sidebar, find the **"Actions"** library
- In the search box, type "**shell**"
- Find **"Run Shell Script"** action
- **Drag it** to the large empty area on the right

**4. Configure the Shell Script**

You should now see a box with script settings. Configure it:

- **Shell:** Select `/bin/bash` from the dropdown (should be default)
- **Pass input:** Select `as arguments` from dropdown

In the large text box, **delete** any existing text and paste this:

```bash
#!/bin/bash
/opt/homebrew/bin/python3 /Users/iamtheking/Desktop/vcard_merger-3.py
```

**5. Save the Application**
- Click `File → Save` (or press `Cmd + S`)
- **Save As:** Type `VCard Merger`
- **Where:** Navigate to your **Applications** folder
  - Important: Save to `/Applications` so it appears in Launchpad
- Click "Save"

**✅ Done!** You now have a working macOS app.

---

#### Part B: Add to Dock (Optional but Recommended)

**1. Open Applications Folder**
- Open Finder
- Press `Cmd + Shift + A` (or Go → Applications)
- Find "VCard Merger" app

**2. Add to Dock**
- **Drag** the "VCard Merger" icon to your Dock
- Place it wherever you want in the Dock
- Release to drop it

**✅ Done!** Click the icon in your Dock anytime to launch the app.

---

#### Part C: Add Custom Icon (Optional)

Want a nice icon? Here's how:

**1. Find or Create an Icon**

Option A - Use an existing icon:
- Find any image you like (contacts, merge, etc.)
- PNG or JPG works fine

Option B - Get a professional icon:
- Search "contacts icon png free" on Google
- Or use SF Symbols (Mac has built-in icons)

**2. Apply the Icon**

**Method 1: From Image File**
- Find your icon image file
- Open it in Preview (double-click)
- Press `Cmd + A` (select all)
- Press `Cmd + C` (copy)
- In Finder, find your "VCard Merger.app" in Applications
- Right-click → "Get Info" (or press `Cmd + I`)
- Click the **small icon** in the top-left corner of the Info window (it highlights blue)
- Press `Cmd + V` (paste)
- Close the Info window
- Your app now has a custom icon!

**Method 2: From Another App**
- Find any app with an icon you like
- Right-click → Get Info
- Click the small icon in top-left
- Press `Cmd + C` (copy)
- Then paste it to your VCard Merger app (follow steps above)

---

### Testing Your New App

**1. Launch from Applications**
- Open Applications folder
- Double-click "VCard Merger"
- The Contact Merger should open!

**2. Launch from Dock**
- Click the icon in your Dock
- App launches instantly

**3. Launch from Spotlight**
- Press `Cmd + Space`
- Type "vcard merger"
- Press Enter

**No Terminal window!** The app launches directly.

---

## Method 2: Simple .command File

**Simpler but less polished** - Opens a Terminal window when you launch it.

### Step-by-Step

**1. Create the File**

Open TextEdit or any text editor and paste this:

```bash
#!/bin/bash
cd /Users/iamtheking/Desktop
/opt/homebrew/bin/python3 vcard_merger-3.py
```

Save as: `VCard-Merger.command` (note the `.command` extension)

**2. Make it Executable**

Open Terminal and run:

```bash
chmod +x /Users/iamtheking/Desktop/VCard-Merger.command
```

**3. Use It**

- Double-click `VCard-Merger.command` to launch
- A Terminal window will open (shows Python output)
- The Contact Merger app will appear

**4. Add to Dock (Optional)**

- Drag the `.command` file to your Dock
- Click to launch anytime

**Pros:**
- Very simple - just one file
- No Automator needed

**Cons:**
- Shows Terminal window
- Less "app-like"
- Not in Applications folder

---

## Method 3: Alias to Desktop (Simplest)

**For quick access without Dock/Applications:**

**1. Create Alias**

Open Terminal and run:

```bash
cat > /Users/iamtheking/Desktop/Launch-VCard-Merger.command << 'EOF'
#!/bin/bash
/opt/homebrew/bin/python3 /Users/iamtheking/Desktop/vcard_merger-3.py
EOF

chmod +x /Users/iamtheking/Desktop/Launch-VCard-Merger.command
```

**2. Use It**

- Double-click the file on your Desktop
- App launches!

---

## Troubleshooting

### "The application cannot be opened"

**Solution:** Check permissions

```bash
chmod +x /Users/iamtheking/Desktop/vcard_merger-3.py
```

### "Python command not found"

**Solution:** Verify Python path

Run this to find your Python:
```bash
which python3
```

Update the script with the correct path (probably `/opt/homebrew/bin/python3` or `/usr/bin/python3`)

### "Permission denied"

**Solution:** Make the script executable

For Automator app:
- The app should work automatically

For .command file:
```bash
chmod +x /path/to/your/file.command
```

### App launches but nothing happens

**Solution:** Check if Python is working

Test in Terminal:
```bash
/opt/homebrew/bin/python3 /Users/iamtheking/Desktop/vcard_merger-3.py
```

If this works but the app doesn't, recreate the Automator app.

---

## Which Method Should I Use?

**For most users:** → **Method 1 (Automator)** ⭐

**Why?**
- True macOS application
- No Terminal window
- Can add custom icon
- Appears in Applications
- Works in Launchpad
- Professional look

**If you want simplicity:** → **Method 2 (.command file)**

**Why?**
- One file, very simple
- No Automator setup
- Works immediately

**For quick testing:** → **Method 3 (Desktop alias)**

**Why?**
- Fastest to set up
- Good for trying it out

---

## Summary

**Recommended: Automator App (Method 1)**

**Quick steps:**
1. Open Automator → New Application
2. Add "Run Shell Script" action
3. Paste: `/opt/homebrew/bin/python3 /Users/iamtheking/Desktop/vcard_merger-3.py`
4. Save to Applications as "VCard Merger"
5. Drag to Dock
6. Done! 🎉

**Result:** Professional macOS app you can launch with one click from your Dock.

---

**Enjoy your one-click Contact Merger!** ✨
