# The event system

Everything in GLF communicates through events. A switch closes and an event is posted; modes
listening for it react and post further events. There are no direct calls between rules.

An event is just a string. `s_LeftOrb1_active`, `mode_base_started`,
`timer_bonus_countdown_complete`, `award_jackpot` — all the same kind of thing.

---

## Posting events

Most of the time you post events declaratively, through an [event player](../players/event-player.md).
When you need to post from your own VBScript — a callback, a table hook — use one of these.

### `DispatchPinEvent evt, kwargs`

The normal one. Queues the event; handlers run on a following frame.

```vbscript
DispatchPinEvent "award_jackpot", Null
```

Queueing rather than calling immediately is what keeps GLF stable: a handler that posts three more
events won't recurse arbitrarily deep inside a single switch hit. The frame budget (the *Glf Frame
Dispatch* option) caps how many are processed per frame; the rest carry over.

### `DispatchQueuePinEvent evt, kwargs`

Same, but handlers may **pause the chain**. A handler returns a `wait_for` kwarg naming another
event; remaining handlers are held until that event arrives.

This is how GLF sequences ball-end: `ball_ending` is a queue event, so a bonus-tally mode can hold
up end-of-ball until its display sequence finishes. You rarely post these yourself — you hook them
with a [queue relay player](../players/queue-relay-player.md).

### `DispatchRelayPinEvent evt, kwargs`

Synchronous, and each handler's return value becomes the next handler's input, so the value is
progressively transformed. Used for veto-style questions — `request_to_start_game` is a relay
event, and any handler can return `False` to block the game from starting.

### `RunAutoFireDispatchPinEvent evt, kwargs`

Synchronous and immediate — bypasses the queue entirely. Reserved for input that must not lag by a
frame: flipper buttons and the coils they drive. Don't reach for it in game logic; the latency win
isn't worth losing the ordering guarantees.

---

## Listening

Declaratively, in a mode:

```vbscript
With .EventPlayer()
    .Add "s_Bumper1_active", Array("score_5000", "play_bumper1_show")
End With
```

Or directly, when you need to run VBScript:

```vbscript
AddPinEventListener "trough_eject", "on_trough_eject", "OnTroughEject", 2000, Null

Function OnTroughEject(args)
    RandomSoundBallRelease swTrough1
    DOF 110, DOFPulse
End Function
```

`AddPinEventListener evt, key, callbackName, priority, args`:

| Parameter | Meaning |
|---|---|
| `evt` | Event name to listen for |
| `key` | Unique handle for this listener — needed to remove it |
| `callbackName` | Name of a global `Function` taking one argument |
| `priority` | Higher runs first |
| `args` | Passed through to the callback as `args(0)` |

The callback receives `Array(your_args, kwargs, event_name)`. For relay and queue events it must
return the kwargs (or a modified copy); for plain events the return value is ignored.

Remove with `RemovePinEventListener evt, key`. Listeners registered by a mode are removed
automatically when the mode stops — this only matters for listeners you register yourself outside
a mode, which live for the whole session.

---

## Event string syntax

Anywhere GLF accepts an event to *listen for*, the string may carry a condition and a priority
offset.

### Conditions — `event{expression}`

The handler only runs if the expression is true.

```vbscript
.Add "s_Plunger1_inactive{current_player.ball_just_started == 1}", Array("new_ball_active")
```

Expressions use [placeholders](placeholders.md) and a small operator set:

| Write | Means |
|---|---|
| `==` | equal |
| `!=` | not equal |
| `&&` | and |
| `\|\|` | or |
| `<` `>` `<=` `>=` | comparisons |

```vbscript
.Add "spinner_hit{current_player.mode_active == 1 && current_player.score > 1000000}", Array("super_spinner")
```

Conditions are compiled once at startup, not parsed per hit, so they're cheap. They're also
evaluated defensively — if an expression errors (typically a variable that doesn't exist yet), it
evaluates to `False` rather than crashing.

### Priority offset — `event.N`

Adds `N` to the mode's priority for this listener only, letting you order handlers *within* a mode.

```vbscript
With .EventPlayer()
    .Add "text_inputted.1{machine.high_score_initials_chars == 3}", Array("text_input_high_score_complete")
End With
```

Higher runs first. Use it when two rules react to the same event and one must observe state before
the other changes it.

### Kwargs on posted events — `event:{key: value}`

When *posting* (the value side of an event player), you can attach data:

```vbscript
.Add "timer_high_score_timeout_complete", Array("text_input_high_score_complete:{text: machine.high_score_initials}")
```

Listeners read it back with the `kwargs.` placeholder:

```vbscript
With .EventName("player_added{kwargs.num == 2}")
```

---

## Kwargs from built-in events

Many GLF devices attach data to the events they post. The useful ones:

| Event | Kwargs |
|---|---|
| `<shot>_hit` | `profile`, `state`, `advancing` |
| `<group>_complete` | `state` |
| `timer_<name>_*` | `ticks`, `ticks_remaining` |
| `timer_<name>_time_added` | `ticks`, `ticks_added`, `ticks_remaining` |
| `multiball_<name>_started` | `balls` |
| `tilt_warning` | `warnings`, `warnings_remaining` |
| `player_added` | `num` |
| `ball_hold_<name>_*`, `multiball_lock_<name>_*` | `balls`, `balls_locked` / `balls_held` |

---

## Ordering and blocking

Handlers for one event run in priority order, highest first. Priority comes from the mode, plus
any `.N` offset. Two handlers at equal priority run in registration order, which you should not
rely on.

Shots additionally **block** an event once they've consumed it: when a shot fires on a switch, it
registers a block so a second shot on the same switch in the same mode won't also fire. A shot
profile with `.Block = True` blocks by shot name rather than shot-plus-profile, which stops the
*same* shot firing twice through different profiles. Blocks clear when the event finishes
dispatching.

---

## Debouncing

GLF ignores a repeated `<switch>_active` for the same switch within 50 ms. This is switch
debounce, not a rate limit on events in general — distinct events are unaffected.

---

## Delays

For one-off timing outside a timer device:

```vbscript
SetDelay "my_delay_name", "MyCallback", Array("some", "args"), 2000   ' ms
RemoveDelay "my_delay_name"
```

Names are unique: setting a delay that already exists replaces it, which makes "restart this
countdown on every hit" a one-liner.

---

## Lifecycle events

The events GLF posts as a game progresses, in order:

| Event | When |
|---|---|
| `reset_complete` | Startup finished (queue event) |
| `request_to_start_game` | Start button released (relay — return `False` to block) |
| `player_added` | A player joins (kwargs: `num`) |
| `game_start` → `game_started` | Game begins |
| `ball_started` | Ball in play; the usual mode start event |
| `trough_eject` | Trough kicked a ball out |
| `ball_drain` | A ball reached the drain (relay) |
| `ball_will_end` | Ball is ending — flippers still live |
| `ball_ending` | Queue event; hold it to delay end-of-ball |
| `ball_ended` | Ball over |
| `next_player` | Rotating to the next player |
| `game_will_end` → `game_ending` → `game_ended` | Game over (`game_ending` is a queue event) |

The distinction that trips people up: `ball_will_end` fires immediately, while `ball_ending` can be
held open by a mode. Use `ball_will_end` to stop gameplay modes, and hook `ball_ending` when you
need to *delay* the end of ball for a bonus sequence.

The full catalogue is in the [event reference](../appendix/event-reference.md).
