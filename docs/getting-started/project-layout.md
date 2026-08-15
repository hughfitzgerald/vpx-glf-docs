# Project layout

A GLF table's logic gets large quickly — a full ruleset is thousands of lines. Keeping it all in
the VPX script editor is workable but unpleasant: no version control, no search across files, and
a single corrupt save loses everything.

The recommended workflow keeps your script as ordinary files on disk, assembled into the `.vpx`
by a watcher process.

## Tools

- **Git** — version control
- **Node.js** — runs the build scripts
- **[vpxtool](https://github.com/francisdb/vpxtool)** by francisdb — extracts and re-injects the
  table script. Add it to your `PATH`.
- An editor with decent VBScript support

## Starting from the example table

```
git clone https://github.com/mpcarr/vpx-example-glf
cd vpx-example-glf/scripts
npm install
npm run rename-project -- MyTableName --git
npm run assemble-vpx
npm run script-watcher
```

`rename-project` renames the table and, with `--git`, resets history so you start with a clean
repo. `assemble-vpx` builds the `.vpx` from the repo contents. `script-watcher` rebuilds the
table script whenever a source file changes — leave it running while you work, and just reload
the table in VPX to pick up changes.

## Suggested structure

```
scripts/
  src/
    glf/
      glf.vbs                 ' the framework, dropped in unmodified
    vpx/                      ' table hardware: flippers, physics, sounds, shadows…
      01_ZCON_Constants.vbs
      04_ZINI_Table_Init.vbs
      ...
    game/                     ' your game logic — this is the interesting part
      _configuration.vbs      ' ConfigureGlfDevices: devices, vars, wiring
      shows/
        general_shows.vbs
        insert_shows.vbs
      modes/
        base.vbs
        attract.vbs
        multiball.vbs
        ...
```

Files are concatenated in path order, so the leading numbers on the `vpx/` files control load
order. The `_configuration.vbs` underscore prefix sorts it first within `game/`.

The split that matters is **`vpx/` versus `game/`**: hardware behaviour versus rules. Table
mechanics — flipper polarity, ball shadows, mechanical sound effects — are borrowed wholesale from
the VPW example table and rarely change. Your rules live in `game/` and change constantly.

## Organising modes

One file per mode, with a `Create<Name>Mode()` sub, all called from `ConfigureGlfDevices`:

```vbscript
' game/modes/multiball.vbs
Sub CreateMultiballMode()
    With CreateGlfMode("multiball", 500)
        ...
    End With
End Sub
```

```vbscript
' game/_configuration.vbs
Sub ConfigureGlfDevices()
    CreateGeneralShows()
    CreateInsertShows()
    CreateSounds()
    CreateSharedShotProfiles()

    CreateBaseMode()
    CreateAttractMode()
    CreateMultiballMode()
    ' …
End Sub
```

Keep the call order: shows and shot profiles before the modes that reference them.

## Shared constants

Colours and light groupings are referenced from many modes. Declaring them once at the top of
`_configuration.vbs` saves a lot of hunting later:

```vbscript
Const JackpotColor   = "00ff6a"
Const MultiballColor = "ffffff"

Dim TargetBankLightNames : TargetBankLightNames = Array("L33","L34","L35","L36","L37","L38")
```

## What not to keep in the repo

The framework writes several artefacts next to the table at runtime. They're regenerated, so
`.gitignore` them:

```
glf_logs/
glf_mpf/
cached-functions.vbs
*_glf.ini
```

`*_glf.ini` holds machine variables and high scores. Deleting it resets them — occasionally handy
during testing.

---

## Next

[The event system →](../concepts/events.md)
