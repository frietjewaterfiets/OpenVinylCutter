# OpenVinylCutter - The First Free Standalone Opensource Vinyl Cutting Software (Currently only for Windows)

<img width="1200" height="567" alt="banner" src="https://github.com/user-attachments/assets/cf964937-6173-4c45-a65d-82bf684fc1f5" />

https://github.com/user-attachments/assets/729d6c25-6f31-41c8-91c5-2c5c943e956d

<img width="1195" height="739" alt="Screenshot_1" src="https://github.com/user-attachments/assets/6eaed7ab-9887-465b-907c-d89c72019699" />


# OpenVinylCutter - Currently only for Windows
This is a small little thing I made because I got tired of paying for a full 1-year subscription, just to cut some vinyl once in a while.
**This app aint perfect. Not even close maybe.** But it already works pretty nice for me, so I still wanted to put it on GitHub.
I hope people that some smart-than-me vinyl cutters, maybe want to test this, tweak this, and help make this project better. Because i could'nt find a good free easy open-sourced software. 
And i can't be the only one. So lets create one together for the community! I tested this mostly with my Secabo S60, but maybe it can also work with other cutters that understand HPGL.

## Before you start
You do need Python for this app.
Make sure Python 3 is installed on your system before running `start_app.bat`.

You can download Python here:
[https://www.python.org/downloads/](https://www.python.org/downloads/)

When installing Python on Windows, make sure to enable:
- `Add Python to PATH`
  
If Python is not installed, `start_app.bat` will not work.

## Installation
Installation should be pretty simple:

1. Download this repository
2. Extract it
3. Run `start_app.bat`

That should install what is needed and start the app.

## What it can do right now
- load SVG files
- drag and drop SVG files into the app
- show a preview
- keep shapes at real size by default
- send the cut to your plotter
- use a peel box around the design
- feed the material forward after cutting
- save HPGL files if you want

## Connection types
There are 2 main ways this app can talk to a cutter.

### `Serial (COM)`
Use this when your cutter shows up as a real COM port in Windows.
DISCLAIMER: I lost my serial cable, sooo i haven't even tested that yet. So that's why there's the other options below. 

### `Windows USB/Printer`
Use this when your cutter is connected by USB, but it does not show up as a normal COM port.
Some cutters show up more like a printer in Windows. In that case this app can send raw HPGL through the Windows print system.

### `Create USB Queue`
This button is there for that second case.
It creates a Windows printer queue on a `USB00x` port, so OpenVinylCutter can send the cut job there.

So in simple words:
- `Serial (COM)` = send straight to a COM port
- `Windows USB/Printer` = send through a Windows printer queue

## Drag and drop
You can just drag an `.svg` file into the app.
If that does not work for some reason, the normal `Load SVG` button still works too.

## Cut settings
These are the main settings in the app.

### `Width (mm)` and `Height (mm)`
This is the material size area the app uses when fitting things.
If auto fit is off, these matter less.

### `Peel box offset (mm)`
This is the space between your design and the outside box.
So if you want a little more room around the design for peeling, increase this a bit.

### `HPGL units/mm`
This one is pretty important.
If the shape is correct, but the size is wrong in real life, this is usually the setting you need.

- too big = lower it
- too small = raise it

`40` is a decent starting point, but not always perfect for every machine.

### `Try this to get your needed HPGL Units/mm:`
Make a little svg file of a simple 50x50mm square, send it to your cutter, and then measure the outcome.

**Then use this formula: **
`new HPGL value = current value x (wanted size / measured size).`

For example: In my case, my cutout square turned out to be around 37x37mm, instead of 50x50mm.
So i did: 40 x (50 / 37) = 54.05. I tried putting in 54.1, and man oh man, Hallaluja, it worked perfectly.

### `Overcut (mm)`
This makes the cut go a tiny bit past the end point.
Useful for small shapes, so corners or closed paths cut a bit cleaner.

### `Feed after cut (mm)`
This tells the app to move the material a bit after cutting is done.
That way the cut is easier to grab and peel.

### `Feed axis`
Not every cutter uses the same axis for feeding the material.
If the machine moves the wrong part, try changing this.

### `Feed direction`
If the feed goes the wrong way, switch this.

## Other options

### `Mirror horizontally`
Useful for things like heat transfer vinyl.

### `Rotate 90 degrees`
Useful when your material is loaded in a different direction than the artwork.

### `Fit to material width automatically`
When this is off, the app tries to keep the real size of the shape itself.
That means empty SVG space is ignored.

### `Cut peel box`
This cuts the outside rectangle around the design as the last cut.
That makes it easier to peel or separate the piece.

### `Flip plotter X axis` and `Flip plotter Y axis`
These are there because some plotters do not move the same way as the preview shows.
So if the preview looks right but the cutter cuts mirrored or flipped, try these.

## Cool little addon: Illustrator script
There is also an Illustrator script included:
- `Illustrator Script/export_to_OpenVinylCutter.jsx`

If you want it to always show up inside Illustrator, place the script here:
`/Adobe/Adobe Illustrator (YEAR)/Presets/(Language)/Scripts`

Then restart Illustrator.

After that it should show up under:
`File > Scripts`

What it does:

- exports your current selection to SVG
- opens OpenVinylCutter
- loads that exported SVG into the app

Important:
- on first run, the script will ask for your OpenVinylCutter folder location
- after that it remembers the location (hopefully)
- if you move the OpenVinylCutter folder later, it should ask again automatically

## First time using it
I would do it like this:

- Start the app
- Load a very simple SVG first
- Keep auto fit off
- Connect your cutter
- Click `Refresh`
- Try and find the correct connection type (Using USB? try selecting different USB ports. For example, select USB001, then select 'Create USB Queue' and then select 'Send to Plotter'. Do this untill you find the right port! (Yes, i know, very professional software).
- If the size is wrong, change `HPGL units/mm` (Read this section somewhere above: 'Try this to get your needed HPGL Units/mm').
- If it cuts mirrored, try the axis flip options
- If it feeds wrong after cutting, change feed axis or direction

## Important to know

- this app is still early
- it is not a perfect SignCut replacement yet
- it only outputs HPGL right now
- some things still need manual tweaking depending on your cutter
- Windows USB/Printer support depends on printer queues, not on fancy official cutter drivers

## Why I made this
Pretty simple.
I just wanted to cut vinyl without having to pay for software again that I barely use.
So this project is mostly for people like me.
If you know more than me and want to improve this little piece of software, then please do so! But you CANNOT sell this. We need all smart vinyl cutters to come together and create THE BEST FREE SOFTWARE!
