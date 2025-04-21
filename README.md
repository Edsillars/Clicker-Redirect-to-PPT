# Clicker Override

## Problem
When showing a presentation using an extended screen and using a remote clicker, the presentation window has to be focused for the clicker to work. This makes it impossible to use the computer and show a presentation at the same time. 

## Solution
An AutoHotkey script to enable usage of remote presentation clickers in background windows. The script brings the presentation back into focus when videos are played as they will not play when powerpoint is a background window.

## How it works
Remote clickers worth by sending Page Up and Page Down or sometimes Left and Right keystrokes to PowerPoint. This AutoHotkey script intercepts those keys and sends them to the PowerPoint Presenter View Window. You can select which keys you would like to redirect. There is also a text box provided where you can type which screen videos appear on, this will bring PowerPoint back to the foreground (focus) for these slides which will allow the video to play, as it will not play in the background.

## Screenshot
![ClickerOverrideGui Screenshot](screenshot.png)


