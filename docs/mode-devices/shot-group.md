# Shot group

Groups [shots](shot.md) so you can act on them together and detect when they all reach the same
state.

```vbscript
With .ShotGroups("target_bank")
    .Shots = Array("target1", "target2", "target3", "target4")
End With
```

## Events posted

When every shot in the group reaches the same state, the group posts:

```
target_bank_complete            (kwargs: state)
target_bank_<statename>_complete
```

So a group of shots on an `unlit`/`lit` profile posts `target_bank_lit_complete` when the last one
is lit — the "you finished the bank" hook:

```vbscript
With .EventPlayer()
    .Add "target_bank_lit_complete", Array("light_lock", "play_voc_lock_is_lit")
End With
```

It fires on *any* shared state, including all-unlit after a reset, so condition on the state name
rather than the bare `_complete` event when that matters.

## Configuration

| Property | Meaning |
|---|---|
| `Shots` | Array of shot names in this mode |
| `EnableEvents` / `DisableEvents` | Enable/disable every shot in the group |
| `ResetEvents` | Reset every shot to state 0 |
| `RestartEvents` | Reset and re-enable every shot |
| `RotateEvents` | Rotate states between shots |
| `RotateLeftEvents` | Rotate left |
| `RotateRightEvents` | Rotate right |
| `EnableRotationEvents` | Allow rotation |
| `DisableRotationEvents` | Disallow rotation |

## Rotation

Rotation shifts states between the shots without changing which states exist — the classic
flipper-controlled "move the lit arrow" mechanic:

```vbscript
With .ShotGroups("lane_lights")
    .Shots = Array("lane1", "lane2", "lane3", "lane4")
    .EnableRotationEvents = Array("mode_base_started")
    .RotateLeftEvents  = Array("s_left_flipper_active")
    .RotateRightEvents = Array("s_right_flipper_active")
End With
```

Rotation is off until an `EnableRotationEvents` fires — without it, the rotate events do nothing.

Shots whose current state is listed in the profile's `StateNamesNotToRotate` are held out of the
rotation, so completed shots can stay put while unfinished ones move around them.

## Reading common state

```vbscript
.Add "s_Scoop_active{device.shot_groups.target_bank.common_state == 1}", Array("collect_bank_award")
```

`common_state` gives the shared state index, or empty when the shots disagree.

## Resetting

```vbscript
With .ShotGroups("target_bank")
    .Shots = Array("target1", "target2", "target3")
    .ResetEvents = Array("target_bank_lit_complete")
End With
```

Careful with that pattern: resetting on the group's own completion event re-arms the bank
immediately, which is usually right for a repeatable award and wrong if you wanted it to stay
completed.

## See also

[Shot](shot.md) · [Shot profile](shot-profile.md)
