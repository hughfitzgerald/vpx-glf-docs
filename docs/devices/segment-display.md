# Segment display

An alphanumeric or numeric display built from GLF-controlled lights — the classic 7- or 14-segment
score displays. Each segment is a light, and GLF maps characters onto them.

```vbscript
Dim segment_display_p1
Set segment_display_p1 = (New GlfLightSegmentDisplay)("player1")

segment_display_p1.SegmentType = "14Segment"
segment_display_p1.SegmentSize = 8
segment_display_p1.LightGroup = "p1_seg"
segment_display_p1.UpdateMethod = "stack"
segment_display_p1.DefaultColor = SegmentsColor
segment_display_p1.UseDotsForCommas = True
```

Displays are created in `ConfigureGlfDevices`. Note the constructor syntax — `New` with the name in
parentheses, not a `CreateGlf…` function.

## Configuration

| Property | Default | Meaning |
|---|---|---|
| `SegmentType` | — | `"14Segment"` or `"7Segment"` |
| `SegmentSize` | — | Number of character positions |
| `LightGroup` | — | Light tag holding this display's segment lights |
| `LightGroups` | — | Array of tags, for a display spanning several groups |
| `UpdateMethod` | — | `"stack"` for priority-stacked text entries |
| `DefaultColor` | white | Segment colour |
| `UseDotsForCommas` | `False` | Render commas using the decimal points |
| `IntegratedCommas` | `False` | Display has dedicated comma segments |
| `IntegratedDots` | `False` | Display has dedicated dot segments |
| `DefaultTransitionUpdateHz` | — | Animation rate for transitions |
| `ExternalFlexDmdSegmentIndex` | — | Offset when mirroring to a FlexDMD |
| `ExternalB2SSegmentIndex` | — | Offset when mirroring to a B2S backglass |

## Lights and tags

The segment lights are identified by a tag, set in each light's **BlinkPattern** field in VPX (see
[Lights & shows](../concepts/lights-and-shows.md#tags)). A 14-segment, 8-character display needs
8 × 15 lights, all tagged `p1_seg`, in order.

Ordering is positional — GLF walks the tagged lights in collection order and assigns them to
segments. Getting the order wrong produces scrambled characters, which is the usual symptom of a
mis-tagged display.

## Multi-display groups

`LightGroups` combines several physical displays into one logical one, so text can span them:

```vbscript
Dim segment_display_all
Set segment_display_all = (New GlfLightSegmentDisplay)("all")

segment_display_all.SegmentType = "14Segment"
segment_display_all.SegmentSize = 32
segment_display_all.LightGroups = Array("p1_seg", "p2_seg", "p3_seg", "p4_seg")
segment_display_all.UpdateMethod = "stack"
segment_display_all.DefaultColor = SegmentsColor
segment_display_all.UseDotsForCommas = True
segment_display_all.DefaultTransitionUpdateHz = 10
```

Now `"all"` is a 32-character display across all four player windows — the standard way to show
mode text on an EM-style backglass. `SegmentSize` must equal the total across the groups.

Displays can overlap: the same lights can belong to `player1` and to `all`. The
[display stack](../players/segment-display-player.md#the-display-stack) resolves which text wins by
priority, so a message on `all` overrides the individual scores and restores them when removed.

## Writing to a display

Through the [segment display player](../players/segment-display-player.md):

```vbscript
With .SegmentDisplayPlayer()
    With .EventName("mode_base_started")
        With .Display("player1")
            .Text = "{players[0].score:0>2,}"
        End With
        With .Display("ball")
            .Text = "{current_player.ball:0>2}"
        End With
    End With
End With
```

## Mirroring to FlexDMD or B2S

`ExternalFlexDmdSegmentIndex` and `ExternalB2SSegmentIndex` give the offset at which this display's
characters appear in an external 32-segment array:

```vbscript
segment_display_p1.ExternalFlexDmdSegmentIndex = 0
segment_display_p1.ExternalB2SSegmentIndex = 1
segment_display_p2.ExternalFlexDmdSegmentIndex = 8
segment_display_p2.ExternalB2SSegmentIndex = 9
```

With the *Glf Virtual Segment DMD* table option on, GLF opens a FlexDMD window mirroring the
segments — useful for testing without a physical backglass, and for players running desktop mode.

## A full four-player setup

```vbscript
Dim segment_display_ball
Set segment_display_ball = (New GlfLightSegmentDisplay)("ball")
segment_display_ball.SegmentType = "14Segment"
segment_display_ball.SegmentSize = 2
segment_display_ball.DefaultColor = SegmentsColor
segment_display_ball.LightGroup = "ball_seg"
```

Then one 8-character display per player as above, plus the combined `all` display for mode text.

## See also

[Segment display player](../players/segment-display-player.md) ·
[Lights & shows](../concepts/lights-and-shows.md)
