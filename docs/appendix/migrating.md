# Migrating older tables

This documentation describes the current `main` branch of
[mpcarr/vpx-glf](https://github.com/mpcarr/vpx-glf). Tables built against earlier versions — and
the published example tables — differ in a few places.

Everything here is **additive or internal**. No documented API was removed, so an older table's
configuration remains valid; what follows is what's new and what changed underneath.

---

## New since the published example table

`vpx-example-glf` predates several additions. All are optional.

### `Mode.GameMode`

```vbscript
.GameMode = False
```

Marks a mode as running outside normal ball play. Used by the built-in high score mode. See
[Modes](../concepts/modes.md).

### `MultiballLocks.LockedBallCountingStrategy`

```vbscript
.LockedBallCountingStrategy = "virtual_only"
```

Controls how locked balls are counted. Defaults to `"virtual_only"`, which is the previous
behaviour, so existing configurations are unaffected.

### `BallDevice.EjectEvents`

Previously only `EjectAllEvents` existed. `EjectEvents` ejects a **single** ball:

```vbscript
With CreateGlfBallDevice("lock_saucer")
    .EjectEvents = Array("release_one_ball")
    .EjectAllEvents = Array("release_all_balls")
End With
```

For single-ball devices the two are equivalent.

### Sounds in show steps

A show step can now trigger a sound directly, keeping audio in sync with lighting:

```vbscript
With .AddStep(Null, Null, 0.5)
    .Lights = Array("(lights)|100|ff0000")
    With .Sounds("sfx_explosion")
    End With
End With
```

Previously this needed a separate event and a sound player entry.

---

## Changed behaviour

### The built-in high score mode

Renamed and re-prioritised:

| | Old | Current |
|---|---|---|
| Mode name | `glf_high_scores` | `high_scores` |
| Priority | 80 | 1500 |
| `GameMode` | (not set) | `False` |

If you referenced `mode_glf_high_scores_started` anywhere, update it to
`mode_high_scores_started`. The much higher priority means the built-in mode's handlers now run
ahead of most table modes — relevant if you built your own initials entry that assumed otherwise.

See [High score](../mode-devices/high-score.md).

### BCP contexts drop the `mode_` prefix

Slides, widgets and sounds sent to a media controller now report their context as `base` rather
than `mode_base`. If your Godot/GMC project matches on context names, update them.

### Events to a media controller carry kwargs

`Glf_BcpSendEvent` now forwards event kwargs, and `PlaySlide` accepts them. A media controller can
read data off the event instead of needing a separate variable push.

### Lightmap discovery

Lightmaps were matched by looking for `_<lightname>_` as a substring of the element name. The
current version builds a token-span index instead: element names are split on underscores and
matched against light names *and* light tags, including multi-token spans.

The practical effect is that existing names keep working and tag-based lightmap naming
(`lm_GI_left` following every light tagged `GI`) now resolves too. If a lightmap stopped following
its light, check the name splits into underscore-delimited tokens that include the light name.

### Expression compilation

Compiled placeholder and condition functions are written to `cached-functions.vbs` and loaded from
there, rather than being executed inline as they are built. This is a startup-time improvement,
and it gives you a readable dump of every generated expression — see [Debugging](debugging.md).

---

## Not in the framework

If you are reading a table's bundled `glf.vbs` as a reference, be aware it may carry local
additions that are not part of GLF. The `darkchaos` table, for instance, includes Godot scene
export helpers (`ToTres`, `GetFileNameFromPath`) that don't exist upstream.

Configuration copied from a table works only if the framework version you're running actually has
the property. When in doubt, the GLF repository is the authority.

---

## Checking your version

There is no version constant. To tell whether a bundled `vpx-glf.vbs` has a given feature, search
it for the property name — for example `Public Property Let GameMode`. If the setter isn't there,
the feature isn't either.
