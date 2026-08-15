# Variables

GLF has two stores: **player variables**, reset each game and tracked per player, and **machine
variables**, shared across players and optionally persisted to disk.

---

## Player variables

### Declaring

Declare initial values in `ConfigureGlfDevices`. Every player gets them at game start.

```vbscript
Glf_SetInitialPlayerVar "ball_just_started", 1
Glf_SetInitialPlayerVar "scoring_multiplier", 1
Glf_SetInitialPlayerVar "ramps_made", 0
```

Undeclared variables still work — they read as `False` until first written — but declaring them
gives a predictable starting value and documents the ruleset in one place.

### Built-ins

| Variable | Meaning |
|---|---|
| `score` | Current score |
| `ball` | Ball number for this player |
| `extra_balls` | Pending extra balls |
| `initials` | Entered initials |
| `number` | Player number, 1-based |
| `tilt_warnings` | Warnings so far (once a tilt device exists) |

GLF also stores internal state as player variables: `shot_<name>` for shot states,
`state_machine_<name>`, `counter_<name>_state`, `extra_ball_<name>_awarded`. Treat those as
read-only, but they're readable, which is occasionally useful.

### Writing

Normally through the [variable player](../players/variable-player.md):

```vbscript
With .VariablePlayer()
    With .EventName("ramp_made")
        With .Variable("score")
            .Action = "add"
            .Int = 25000
        End With
        With .Variable("ramps_made")
            .Action = "add"
            .Int = 1
        End With
    End With
End With
```

From your own script:

```vbscript
SetPlayerState "ramps_made", GetPlayerState("ramps_made") + 1
value = GetPlayerStateForPlayer(0, "score")     ' zero-indexed
```

### Reading

In configs, via [placeholders](placeholders.md):

```vbscript
.Int = "{current_player.jackpot_value}"
.Add "s_Ramp_active{current_player.ramps_made >= 3}", Array("ramps_complete")
.Text = "{players[0].score:0>2,}"
```

### Change events

Writing a player variable can drive logic directly, without an explicit event:

```vbscript
AddPlayerStateEventListener "ramps_made", "my_ramp_watcher", -1, "OnRampsChanged", 100, Null

Function OnRampsChanged(args)
    Dim data : data = args(1)     ' Array(name, newValue, previousValue)
    If data(1) >= 3 Then DispatchPinEvent "ramps_complete", Null
End Function
```

The `-1` is the player index to watch; `-1` means the current player. Remove with
`RemovePlayerStateEventListener`.

These handlers fire on a short delay after the write, and only when the value actually changes —
setting a variable to the value it already holds is a no-op. That matters if you're counting
writes rather than reacting to values.

For most rules an event player condition is simpler and easier to follow than a state listener.

---

## Machine variables

Shared across players; survive between games; optionally written to `<cGameName>_glf.ini`.

### Declaring

```vbscript
With CreateMachineVar("bottom_ball_locked")
    .InitialValue = 0
    .ValueType = "int"        ' "int" or "string"
    .Persist = False          ' True to save to disk
End With
```

`Persist = True` writes the value at table exit and restores it at startup. Use it for credits,
audits and anything that should survive a restart. `Persist = False` still shares the value across
players within a session, which is the common case for things like "is the lock occupied".

### Using

```vbscript
' read
.Add "s_Lock_active{machine.bottom_ball_locked == 0}", Array("lock_ball")

' write
With .VariablePlayer()
    With .EventName("lock_ball")
        With .Variable("bottom_ball_locked")
            .Action = "set_machine"
            .Int = 1
        End With
    End With
End With
```

Note the action names: player variables use `add` / `set`, machine variables use `add_machine` /
`set_machine`. Using the wrong pair is a common and quiet mistake — the write goes to the wrong
store and the condition you expected to fire never does.

GLF predefines `player1_score` … `player4_score` as persisted machine variables, updated at game
end, which is what an attract-mode display typically reads.

---

## Choosing between them

| Use a player variable when | Use a machine variable when |
|---|---|
| Progress belongs to one player | State belongs to the machine |
| It should reset each game | It should persist across games |
| Each player tracks it separately | All players see the same value |

The awkward middle case is a physical lock: the ball is in the saucer regardless of whose turn it
is, so the *occupancy* is a machine variable, while *credit* for the lock is a player variable.
Both example tables handle it that way.

---

## Persistence and the ini file

Persisted machine variables and high scores are stored in `<cGameName>_glf.ini`, written on table
exit. Delete it to reset everything — useful during development, and worth remembering before you
chase a "high scores won't clear" bug.
