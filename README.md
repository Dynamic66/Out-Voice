# Out-Voice

### Copy text and let it be read out to you

  ![image](https://github.com/user-attachments/assets/61136e0d-e3c6-42a7-b392-8ce7b3079fac)

## Past this into a powershell terminal to immediately use Out-Voice 
  ```
  Invoke-RestMethod https://raw.githubusercontent.com/Dynamic66/Out-Voice/refs/heads/main/Out-Voice.ps1 | Invoke-Expression
  ```

## Usage
Top controls:
- AutoCopy: Immediately copies clipboard text into the textbox below
- AutoPlay: Immediately reads content of the textbox below. Requires AutoCopy to be enabled 
- AlwaysOnTop: Keeps the window above all other windows

Textbox: Select a part of the text and click play to read it out. 

Bottom controls:
- Combo box: Select an installed voice (can be installed over Windows Setting > Language > Preferred languages)
- Play/Pause button: Play/Pause/Resume Asynchronous reading
- Stop button: stops Asynchronous reading
- Volume slider: Change the Volume. Cant be changed while reading is in progress
- Numeric Up Down: Change the reading speed from -10 to 10. Cant be changed while reading is in progress
