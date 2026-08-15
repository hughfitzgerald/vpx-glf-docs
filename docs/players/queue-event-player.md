# Queue event player

Posts **queue events** — events whose handlers can pause the chain until something else finishes.

```vbscript
With .QueueEventPlayer()
    .Add "start_bonus_sequence", Array("bonus_ready")
End With
```

The API mirrors the [event player](event-player.md): `.Add key, values`, with conditions and
priority offsets on the key.

## When you need it

A queue event lets a handler say "hold on, wait for *this* before continuing". GLF uses it
internally for the ball-end sequence: `ball_ending` is a queue event, so a bonus mode can hold it
open while it counts down, and only when the mode stops does end-of-ball proceed.

You post a queue event yourself when you're building the same shape: a sequence of steps where a
later step must not begin until an earlier one signals completion.

```vbscript
With .QueueEventPlayer()
    .Add "mode_wizard_started", Array("wizard_intro")
End With
```

Whatever listens for `wizard_intro` can now delay everything downstream — a mode started by it
with `UseWaitQueue = True`, or a [queue relay player](queue-relay-player.md) entry.

## Holding the queue

Two ways to hold a queue event open:

**A mode with `UseWaitQueue`** — the queue waits until the mode stops:

```vbscript
With CreateGlfMode("bonus", 1000)
    .StartEvents = Array("ball_ending")
    .StopEvents  = Array("bonus_complete")
    .UseWaitQueue = True
End With
```

**A queue relay player** — posts an event and waits for a named reply:

```vbscript
With .QueueRelayPlayer()
    With .EventName("ball_ending")
        .Post = "start_bonus"
        .WaitFor = "bonus_complete"
    End With
End With
```

## Notes

- Values are plain event names — this player does not accept the `event:{kwargs}` form.
- If nothing holds the queue, a queue event behaves like an ordinary one.
- Forgetting to post the `WaitFor` event stalls the sequence indefinitely. When end-of-ball hangs,
  an unanswered wait is the first thing to check.

## See also

[Event player](event-player.md) · [Queue relay player](queue-relay-player.md) ·
[Modes](../concepts/modes.md)
