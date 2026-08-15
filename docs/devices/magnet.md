# Magnet

Grabs, holds, releases and flings a ball.

```vbscript
With CreateGlfMagnet("mag1")
    .GrabSwitch = "s_ST4"
    .EnableEvents  = Array("ball_started")
    .DisableEvents = Array("ball_ended")
    .FlingBallEvents = Array("magnet_mag1_grabbed_ball")
    .GrabTime = 500
    .FlingDropTime = 100
    .FlingRegrabTime = 50
    .ActionCallback = "GrabMagnetAction"
End With
```

## Configuration

| Property | Meaning |
|---|---|
| `ActionCallback` | Function that energises the magnet |
| `GrabSwitch` | Switch that triggers an automatic grab |
| `GrabBallEvents` | Events that grab a ball |
| `ReleaseBallEvents` | Events that release it gently |
| `FlingBallEvents` | Events that fling it |
| `ResetEvents` | Events that reset the magnet |
| `GrabTime` | Ms to energise before the ball counts as grabbed |
| `ReleaseTime` | Ms for a release to complete |
| `FlingDropTime` | Ms the magnet is off during a fling |
| `FlingRegrabTime` | Ms before the magnet re-energises after a fling |
| `EnableEvents` / `DisableEvents` | Enable/disable the device |

Times accept [placeholders](../concepts/placeholders.md).

## The action callback

Receives 1 to energise, 0 to release:

```vbscript
Sub GrabMagnetAction(enabled)
    If enabled = 1 Then
        GrabMag.MagnetOn = True
    Else
        GrabMag.MagnetOn = False
    End If
End Sub
```

The magnet itself is a standard `cvpmMagnet` from the VPX core, set up in `Table1_Init`:

```vbscript
Set GrabMag = New cvpmMagnet
With GrabMag
    .InitMagnet GrabMagnet, 30
    .GrabCenter = False
    .strength = 15
    .CreateEvents "GrabMag"
End With
```

## Events posted

| Event | When |
|---|---|
| `magnet_<name>_grabbing_ball` | Grab started |
| `magnet_<name>_grabbed_ball` | Grab complete, after `GrabTime` |
| `magnet_<name>_releasing_ball` | Release started |
| `magnet_<name>_released_ball` | Release complete |
| `magnet_<name>_flinging_ball` | Fling started |
| `magnet_<name>_flinged_ball` | Fling complete |

## Fling versus release

**Release** turns the magnet off and lets the ball drop.

**Fling** cuts power briefly (`FlingDropTime`), then re-energises (`FlingRegrabTime`) — the ball
falls, is caught again, and is thrown by the sudden pull. Tune the two times against your magnet
strength; small changes make a large difference.

The example configuration flings immediately on grab, chaining the events:

```vbscript
.FlingBallEvents = Array("magnet_mag1_grabbed_ball")
```

Grab completes, which triggers the fling. For a magnet that holds the ball while something happens
first, fling on a different event instead:

```vbscript
With CreateGlfMagnet("mag1")
    .GrabSwitch = "s_ST4"
    .GrabTime = 500
    .FlingBallEvents = Array("magnet_award_complete")
    .ActionCallback = "GrabMagnetAction"
End With

With .EventPlayer()
    .Add "magnet_mag1_grabbed_ball", Array("start_magnet_award")
End With
```

## Notes

- A grab is ignored while a release or fling is in progress.
- Disable the magnet at ball end, or a held ball never drains.

## See also

[Diverter](diverter.md) · [Ball device](ball-device.md)
