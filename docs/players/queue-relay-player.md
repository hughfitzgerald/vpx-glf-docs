# Queue relay player

Hooks a queue event, posts an event of its own, and holds the queue until a named event answers.

```vbscript
With .QueueRelayPlayer()
    With .EventName("ball_ending")
        .Post = "start_bonus_tally"
        .WaitFor = "bonus_tally_complete"
    End With
End With
```

Read it as: *when `ball_ending` arrives, post `start_bonus_tally` and don't let `ball_ending`
finish until `bonus_tally_complete` fires.*

| Property | Meaning |
|---|---|
| `Post` | Event to post immediately |
| `WaitFor` | Event that releases the queue |

The `EventName` key accepts conditions and priority offsets like any other listener.

## Why this exists

End-of-ball is the classic case. `ball_ending` is a queue event, and several things may want to
run before the ball actually ends: bonus tally, a "shoot again" award animation, a mode wrap-up.
Each can insert itself with a relay entry, and they run in priority order without knowing about
one another.

```vbscript
' In the bonus mode
With .QueueRelayPlayer()
    With .EventName("ball_ending")
        .Post = "start_bonus"
        .WaitFor = "bonus_complete"
    End With
End With

With .EventPlayer()
    .Add "start_bonus", Array("play_bonus_show", "start_bonus_timer")
    .Add "timer_bonus_complete", Array("bonus_complete")
End With
```

## Conditional holds

Skip the hold when it isn't wanted — a tilted ball shouldn't sit through a bonus count:

```vbscript
With .EventName("ball_ending{game.tilted == False}")
    .Post = "start_bonus"
    .WaitFor = "bonus_complete"
End With
```

When the condition is false the entry does nothing and the queue proceeds.

## Notes

- Only useful against events posted as **queue events** — `ball_ending`, `game_ending`,
  `mode_<name>_starting` / `_stopping`, `reset_complete`, and anything from a
  [queue event player](queue-event-player.md).
- `WaitFor` must eventually be posted. Give the mode a timer as a backstop if the releasing event
  depends on player action.
- A mode with `UseWaitQueue = True` is often simpler when the thing you're waiting for is
  "this whole mode finishing".

## See also

[Queue event player](queue-event-player.md) · [The event system](../concepts/events.md)
