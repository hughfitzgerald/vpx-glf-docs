# Sound player

Plays and stops sounds defined with `CreateGlfSound`.

```vbscript
With .SoundPlayer()
    With .EventName("play_sfx_jackpot")
        .Key = "key_sfx_jackpot"
        .Sound = "sfx_jackpot"
    End With
End With
```

| Property | Default | Meaning |
|---|---|---|
| `Sound` | — | Name of a sound created with `CreateGlfSound` |
| `Key` | — | Handle for stopping it later |
| `Action` | `"play"` | `play` or `stop` |
| `Volume` | sound's own | Overrides the sound's and the bus's volume |
| `Loops` | sound's own | 0 once, -1 forever, N repeats |

## Music: start and stop

Long-running audio needs an explicit stop, usually tied to the mode:

```vbscript
With .SoundPlayer()
    With .EventName("play_mus_ambient_loop")
        .Key = "key_mus_ambient_loop"
        .Sound = "mus_ambient_loop"
    End With
    With .EventName("stop_mus_ambient_loop")
        .Key = "key_mus_ambient_loop"
        .Sound = "mus_ambient_loop"
        .Action = "stop"
    End With
End With

With .EventPlayer()
    .Add "mode_base_started",  Array("play_mus_ambient_loop")
    .Add "mode_base_stopping", Array("stop_mus_ambient_loop")
End With
```

## Buses

Every sound belongs to a bus, which caps polyphony and sets a default volume. Buses are where you
keep effects from drowning out callouts — see [Sound system](../devices/sound-system.md).

```vbscript
With CreateGlfSoundBus("voc")
    .SimultaneousSounds = 2
    .Volume = 1
End With
```

## Notes

- Sounds must be defined before the modes that reference them.
- Playing a sound that's already playing restarts it.
- A sound's `Duration` matters: GLF uses it to know when playback ended, so a wrong duration means
  the sound is never marked stopped and `EventsWhenStopped` won't fire.

## See also

[Sound system](../devices/sound-system.md)
