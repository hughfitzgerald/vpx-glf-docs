# Sound system

Sounds are declared once, assigned to a bus, and played by the
[sound player](../players/sound-player.md).

## Buses

A bus groups sounds, caps how many play at once, and sets a default volume.

```vbscript
With CreateGlfSoundBus("mus")
    .SimultaneousSounds = 4
    .Volume = 1
End With

With CreateGlfSoundBus("sfx")
    .SimultaneousSounds = 8
    .Volume = 0.5
End With

With CreateGlfSoundBus("voc")
    .SimultaneousSounds = 2
    .Volume = 1
End With
```

| Property | Default | Meaning |
|---|---|---|
| `SimultaneousSounds` | 8 | Maximum concurrent sounds on this bus |
| `Volume` | 0.5 | Default volume, 0–1 |
| `SystemType` | — | Set to `"bcp"` to route playback to a media controller |

The music / effects / callouts split is conventional and worth keeping: it lets you hold effects
below callouts so speech is never buried, and the low `SimultaneousSounds` on `voc` stops callouts
stacking into noise.

Setting `SystemType = "bcp"` hands playback to an external media controller rather than VPX —
useful when a Godot project owns the audio.

## Sounds

```vbscript
With CreateGlfSound("sfx_jackpot")
    .File = "sfx_jackpot"      ' name in the VPX Sound Manager
    .Bus = "sfx"
    .Duration = 1.386          ' seconds
End With
```

| Property | Meaning |
|---|---|
| `File` | Name of the sound in the VPX Sound Manager |
| `Bus` | Bus to play on |
| `Duration` | Length in seconds |
| `Volume` | Override the bus volume |
| `Loops` | 0 once, -1 forever, N repeats |
| `Priority` | Priority within the bus |
| `MaxQueueTime` | How long a queued sound waits |
| `EventsWhenStopped` | Events posted when the sound finishes |

**`Duration` matters.** GLF uses it to know when playback ended and free the bus slot. If it's
wrong, the sound is never marked stopped: the slot stays occupied, `EventsWhenStopped` never fires,
and eventually the bus refuses new sounds. Read the real durations off the files.

## Helper subs

Declaring dozens of sounds is repetitive, so wrap it:

```vbscript
Sub AddMusic(Name, Duration, Loops)
    With CreateGlfSound(Name)
        .File = Name
        .Bus = "mus"
        .Loops = Loops
        .Duration = Duration
    End With
End Sub

Sub AddSoundEffect(Name, Duration)
    With CreateGlfSound(Name)
        .File = Name
        .Bus = "sfx"
        .Duration = Duration
    End With
End Sub

Sub AddCallout(Name, Duration)
    With CreateGlfSound(Name)
        .File = Name
        .Bus = "voc"
        .Duration = Duration
    End With
End Sub
```

```vbscript
Sub CreateSounds()
    AddMusic "mus_ambient_loop", 235.848, -1
    AddMusic "mus_multiball_loop", 38.792, -1

    AddSoundEffect "sfx_jackpot", 1.386
    AddSoundEffect "sfx_drop_target", 0.961

    AddCallout "voc_multiball_ready", 1.442
    AddCallout "voc_jackpot1", 0.995
End Sub
```

Call `CreateSounds()` from `ConfigureGlfDevices`, before any mode that references a sound.

## Positioning

Sounds played through GLF are not positional. Place them at the backglass in the VPX Sound Manager
so they play centred — playfield-positional effects (ball rolling, flipper mechanics) should stay
with the standard VPX sound routines rather than going through GLF.

## Chaining on completion

```vbscript
With CreateGlfSound("voc_intro")
    .File = "voc_intro"
    .Bus = "voc"
    .Duration = 3.2
    .EventsWhenStopped = Array("intro_callout_finished")
End With
```

This depends on `Duration` being right.

## See also

[Sound player](../players/sound-player.md)
