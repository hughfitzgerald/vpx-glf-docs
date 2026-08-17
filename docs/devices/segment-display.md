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
characters appear in an external segment array. Both are independent of your playfield/backglass
lights — the display still drives those normally through `LightGroup` — and each target gets its
own separate push, so a display can mirror to one, both, or neither.

```vbscript
segment_display_p1.ExternalFlexDmdSegmentIndex = 0
segment_display_p1.ExternalB2SSegmentIndex = 0
segment_display_p2.ExternalFlexDmdSegmentIndex = 8
segment_display_p2.ExternalB2SSegmentIndex = 8
```

Offset by however many characters the previous display used, same as `player1` at 0 and `player2`
at 8 above.

### FlexDMD

With the *Glf Virtual Segment DMD* table option on, GLF opens a FlexDMD window mirroring the
segments — useful for testing without a physical backglass, and for players running desktop mode.

### B2S

`ExternalB2SSegmentIndex` pushes to a real B2S backglass over the standard `controller` COM
object, via `controller.B2SSetLED`. Two things about it are easy to miss:

**It only works for 14-segment displays.** GLF's B2S output is wired into the `14Segment` branch
of its update loop only; there's no equivalent for `SegmentType = "7Segment"`. A 7-segment display
will still drive its playfield lights correctly, but nothing will reach B2S.

**It's a separate encoding, not a color mirror.** GLF translates each character to a segment
bitmask matching Herweh's DesignerB2S 15-segment LED format and sends that value directly — it
does not forward `DefaultColor` or anything else about how the display looks in VPX. Colour on the
B2S side comes entirely from your B2S backglass configuration.

To use it:

1. Your table needs a working `controller` object — the usual case if the table already uses B2S.
2. In your B2S designer file, add a 15-segment LED bank sized for the characters you want to
   mirror, and note its starting LED index.
3. Set `ExternalB2SSegmentIndex` to that starting index.

```vbscript
segment_display_p1.SegmentType = "14Segment"      ' required — 7Segment has no B2S output
segment_display_p1.SegmentSize = 8
segment_display_p1.LightGroup = "p1_seg"
segment_display_p1.ExternalB2SSegmentIndex = 0
```

Nothing else is needed — text you send through the
[segment display player](../players/segment-display-player.md) reaches B2S automatically once the
index is set, no separate wiring required.

If characters aren't appearing on the backglass, check first that `SegmentType` is `"14Segment"`
and that the index matches your B2S designer file. The B2S call is wrapped in `On Error Resume
Next`, so a missing or misconfigured `controller` fails silently rather than raising an error —
there's no exception to catch to point you at the problem.

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