# Lights & shows

GLF drives lights through **shows** — timed sequences of colour changes — layered on a per-light
priority stack so that several modes can address the same insert without fighting.

---

## Lights

Any light in the `glf_lights` collection is under GLF's control. It's addressed by its VPX name:

```vbscript
With .Lights("L12")
    .Color = "ff8800"
End With
```

Colours are six-digit hex, no `#`.

### Tags

A light's **BlinkPattern** field in VPX is repurposed as a comma-separated tag list. Any tag can
then be used wherever a light name can, addressing every light carrying it:

```
BlinkPattern:  GI, playfield_left
```

```vbscript
With .Lights("GI")        ' every light tagged GI
    .Color = GIColor3000k
    .Fade = 300
End With
```

Tags are how you write "turn the GI amber" once instead of twenty times. Conventional groupings:
`GI`, per-bank tags like `ball_seg` / `p1_seg` for segment displays, and per-feature tags for
inserts that always animate together.

### Lightmaps

If you use lightmaps (baked bloom primitives/flashers), GLF wires them automatically at startup by
matching element names. A primitive or flasher whose name contains the light's name as an
underscore-delimited token — `lm_L12_bloom` for light `L12` — gets its colour driven along with the
light. Tags work the same way: `lm_GI_left` follows every light tagged `GI`.

This means the naming convention is the whole configuration; there's nothing to declare. If a
lightmap isn't following its light, the name is nearly always the reason.

---

## Shows

A show is a list of steps. Each step sets some lights and holds for a duration.

```vbscript
With CreateGlfShow("flicker_color")
    With .AddStep(Null, Null, 0.16)
        .Lights = Array("(lights)|37|(color)")
    End With
    With .AddStep(Null, Null, 0.16)
        .Lights = Array("(lights)|100|(color)")
    End With
    With .AddStep(Null, Null, 0.16)
        .Lights = Array("(lights)|24|(color)")
    End With
End With
```

`AddStep(absoluteTime, relativeTime, duration)` — in practice you pass `Null, Null, duration`, with
duration in **seconds**. (The first two arguments express step timing as absolute or relative
offsets instead; the duration form is what the example tables use throughout.)

A duration of **-1** holds the step forever — the idiom for "end in this state":

```vbscript
With CreateGlfShow("flicker_on")
    With .AddStep(Null, Null, 0.16)
        .Lights = Array("(lights)|37|(color)")
    End With
    With .AddStep(Null, Null, -1)          ' settle here and stay
        .Lights = Array("(lights)|100|(color)")
    End With
End With
```

### Light entry syntax

Each string in `.Lights` is pipe-delimited:

```
<light>|<intensity>|<color>|<fade_ms>
```

| Field | Notes |
|---|---|
| light | Light name, tag name, or a `(token)` |
| intensity | 0–100; scales the colour |
| color | Six-digit hex, a `(token)`, or `stop` |
| fade ms | Optional; fades from the current colour over this many milliseconds |

Only the first two fields are required. `stop` in the colour position removes this show's entry
from that light's stack, letting whatever is underneath show through — different from setting
`000000`, which pushes black on top.

```vbscript
.Lights = Array("L12|100|ff0000", "L13|50|00ff00|500", "GI|100|(color)")
```

### Tokens

`(name)` in a show is a placeholder filled in by whoever plays it. That's what makes one show
reusable:

```vbscript
With CreateGlfShow("flash_color_with_fade")
    With .AddStep(Null, Null, 1)
        .Lights = Array("(lights)|100|(color)|(fade)")
    End With
    With .AddStep(Null, Null, 1)
        .Lights = Array("(lights)|100|000000|(fade)")
    End With
End With
```

```vbscript
With .EventName("play_bumper1_show")
    .Show = "flash_color_with_fade"
    .Speed = 15
    .Loops = 0
    With .Tokens()
        .Add "lights", "tBL1"
        .Add "fade", 500
        .Add "color", Bumper1Color
    End With
End With
```

A token in the light position may resolve to a single light *or* a tag — the show expands
accordingly at cache time.

### Built-in shows

Always available:

| Show | Effect | Tokens |
|---|---|---|
| `on` | Lights on at full | `lights` |
| `off` | Lights off | `lights` |
| `flash` | Flash white | `lights` |
| `flash_color` | Flash a colour | `lights`, `color` |
| `led_color` | Hold a colour | `lights`, `color` |
| `fade_led_color` | Fade to a colour and hold | `lights`, `color`, `fade` |

### Other things in a step

Steps aren't limited to lights. A step can also fire DOF, play a nested show, trigger a sound, or
push a slide or widget:

```vbscript
With .AddStep(Null, Null, 0.5)
    .Lights = Array("(lights)|100|ff0000")
    With .DOFEvent(110)
        .Action = "DOF_PULSE"
    End With
    With .Sounds("sfx_explosion")
    End With
    With .Shows("secondary_flash")
        .Speed = 2
    End With
End With
```

Nested shows are stopped automatically when the parent show moves off that step, which makes
composite effects safe to build.

---

## Playing shows

Through the [show player](../players/show-player.md):

```vbscript
With .ShowPlayer()
    With .EventName("jackpot_lit")
        .Key = "key_jackpot_show"
        .Show = "flash_color"
        .Speed = 4
        .Loops = -1
        .Priority = 500
        With .Tokens()
            .Add "lights", "L20"
            .Add "color", JackpotColor
        End With
    End With
    With .EventName("jackpot_collected")
        .Key = "key_jackpot_show"
        .Action = "stop"
    End With
End With
```

| Property | Default | Meaning |
|---|---|---|
| `Show` | — | Show name |
| `Key` | `""` | Handle for stopping it later — **make it unique** |
| `Action` | `"play"` | `play` or `stop` |
| `Speed` | 1 | Divides step durations; higher is faster |
| `Loops` | -1 | -1 loops forever, 0 plays once, N repeats N times |
| `Priority` | 0 | Added to the mode's priority |
| `Tokens` | — | Token values |
| `EventsWhenCompleted` | — | Events posted when a non-looping show ends |

`Key` is the part people get wrong. Two shows sharing a key will interfere; stopping a show
requires quoting the key you started it with. Prefix them per mode (`key_mb_jackpot`) and they
stay unique by construction.

Shows can also be attached to [shot profile](../mode-devices/shot-profile.md) states, which is how
inserts get their look.

---

## The light stack

Every light has a stack of entries, each tagged with the key that pushed it and a priority. The
top entry wins.

- Playing a show pushes an entry per light it touches.
- Stopping it pops that entry; the light reverts to whatever is underneath, or off if nothing is.
- A higher-priority entry takes the light immediately; when it's removed, the lower one shows
  again without needing to be replayed.

Practically, this means you don't have to restore lights by hand. Multiball can flood the
playfield, and when it ends the base mode's lighting is simply there again.

Where it goes wrong is stacking without unstacking — a mode that plays a looping show and never
stops it leaves an entry behind. Stopping the mode won't necessarily clear it if the show was
keyed oddly. Pair every long-lived `play` with a `stop` on `mode_<name>_stopping`.

---

## Fades

A fourth field on a light entry fades from the current colour:

```vbscript
.Lights = Array("(lights)|100|(color)|800")     ' 800 ms fade
```

Or, from the [light player](../players/light-player.md), which sets a static colour rather than
running a sequence:

```vbscript
With .LightPlayer()
    With .EventName("mode_base_started")
        With .Lights("GI")
            .Color = GIColor3000k
            .Fade = 300
        End With
    End With
End With
```

Use the light player for "set these lights to this colour"; use a show for anything with more than
one step. The light player is cheaper — no per-step scheduling.

---

## Performance

Lights are the most expensive thing GLF does, and a busy playfield can cost frames. Two levers:

- **Glf Min Lightmap Update Rate** (table option) throttles how often lightmap colours are pushed,
  trading smoothness for headroom.
- **Fewer, broader tags.** A show addressing a tag that expands to forty lights schedules forty
  updates per step. Splitting rarely-changing lights out of animated groups helps more than
  micro-tuning step durations.
