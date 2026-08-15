# Event reference

Every event GLF posts. `<name>` is the device or mode name you gave it.

Event names are how rules connect, so this doubles as a map of what you can react to.

---

## Game lifecycle

| Event | Type | Kwargs | When |
|---|---|---|---|
| `reset_complete` | queue | | Startup finished |
| `request_to_start_game` | relay | | Start button released; return `False` to block |
| `player_added` | | `num` | A player joined |
| `game_start` | | | Game beginning |
| `game_started` | | | Game begun |
| `ball_started` | | | Ball in play |
| `trough_eject` | | | Trough kicked a ball out |
| `ball_drain` | relay | | A ball reached the drain |
| `ball_will_end` | | | Ball ending; fires immediately |
| `ball_ending` | queue | | Ball ending; can be held open |
| `ball_ended` | | | Ball over |
| `next_player` | | | Rotating to the next player |
| `game_will_end` | | | Game ending |
| `game_ending` | queue | | Game ending; can be held open |
| `game_ended` | | | Game over |
| `glf_game_cancel` | | | Post this to abandon the game |

## Modes

| Event | Type | When |
|---|---|---|
| `mode_<name>_starting` | queue | Mode starting; devices activate |
| `mode_<name>_started` | | Mode started |
| `mode_<name>_stopping` | queue | Mode stopping; devices still live |
| `mode_<name>_stopped` | | Mode stopped |

## Switches

| Event | When |
|---|---|
| `<switch>_active` | Switch closed |
| `<switch>_inactive` | Switch opened |
| `<slingshot>_active` | Slingshot fired |
| `<spinner>_active` | Spinner turned |

## Keys

| Event | Source |
|---|---|
| `s_start_active` / `s_start_inactive` | Start button |
| `s_left_flipper_active` / `_inactive` | Left flipper button |
| `s_right_flipper_active` / `_inactive` | Right flipper button |
| `s_left_staged_flipper_key_active` / `_inactive` | Upper left flipper |
| `s_right_staged_flipper_key_active` / `_inactive` | Upper right flipper |
| `s_plunger_key_active` / `_inactive` | Plunger key |
| `s_lockbar_key_active` / `_inactive` | Lockbar button |
| `s_left_magna_key_active` / `_inactive` | Left MagnaSave |
| `s_right_magna_key_active` / `_inactive` | Right MagnaSave |
| `s_add_credit_key_active` / `_inactive` | Add credit |
| `s_add_credit_key2_active` / `_inactive` | Add credit 2 |
| `s_tilt_warning_active` | Mechanical tilt or accumulated nudge |

---

## Shots

| Event | Kwargs |
|---|---|
| `<shot>_hit` | `profile`, `state`, `advancing` |
| `<shot>_<profile>_hit` | same |
| `<shot>_<profile>_<state>_hit` | same |
| `<shot>_<state>_hit` | same |

`state` is the state *before* advancing.

## Shot groups

| Event | Kwargs |
|---|---|
| `<group>_complete` | `state` |
| `<group>_<state>_complete` | |

## Sequence shots

| Event | When |
|---|---|
| `<name>_hit` | Sequence completed |
| `sequence_shot_<name>_timeout` | Sequence expired |

## Timers

| Event | Kwargs |
|---|---|
| `timer_<name>_started` | `ticks`, `ticks_remaining` |
| `timer_<name>_tick` | `ticks`, `ticks_remaining` |
| `timer_<name>_complete` | `ticks`, `ticks_remaining` |
| `timer_<name>_stopped` | `ticks`, `ticks_remaining` |
| `timer_<name>_paused` | `ticks`, `ticks_remaining` |
| `timer_<name>_time_added` | `ticks`, `ticks_added`, `ticks_remaining` |
| `timer_<name>_time_subtracted` | `ticks`, `ticks_subtracted`, `ticks_remaining` |

## Ball saves

| Event | When |
|---|---|
| `ball_save_<name>_enabled` | Armed |
| `ball_save_<name>_timer_start` | Countdown started |
| `ball_save_<name>_hurry_up` | About to expire |
| `ball_save_<name>_grace_period` | In the grace window |
| `ball_save_<name>_saving_ball` | A ball was saved |

## Multiball

| Event | Kwargs |
|---|---|
| `multiball_<name>_started` | `balls` |
| `multiball_<name>_ended` | |
| `multiball_<name>_ball_lost` | |
| `multiball_<name>_shoot_again` | |
| `multiball_<name>_shoot_again_ended` | |
| `multiball_<name>_hurry_up` | |
| `multiball_<name>_grace_period` | |
| `multiball_<name>_reset_event` | |
| `ball_save_<name>_timer_start` | |
| `ball_save_<name>_add_a_ball_timer_start` | |

## Multiball locks

| Event | Kwargs |
|---|---|
| `multiball_lock_<name>_locked_ball` | `balls_locked` |
| `multiball_lock_<name>_full` | `balls_locked` |

## Ball holds

| Event | Kwargs |
|---|---|
| `ball_hold_<name>_held_ball` | `balls_held` |
| `ball_hold_<name>_full` | `balls_held` |
| `ball_hold_<name>_balls_released` | |

## Extra balls

| Event | When |
|---|---|
| `extra_ball_<name>_awarded` | This device awarded one |
| `extra_ball_awarded` | Any extra ball awarded |

## Tilt

| Event | Kwargs |
|---|---|
| `tilt_warning` | `warnings`, `warnings_remaining` |
| `tilt_warning_<n>` | |
| `tilt` | |
| `tilt_mode_<modename>_clear` | |

## High score

| Event | When |
|---|---|
| `high_score_enter_initials` | Initials needed |
| `high_score_award_display` | Award being shown |
| `<category>_award_display` | Award for a category |
| `<award>_award_display` | Award for a label |
| `high_score_complete` | Entry finished |

---

## Ball devices

| Event | When |
|---|---|
| `balldevice_<name>_ball_entered` | Ball settled |
| `balldevice_<name>_ball_exiting` | Ball leaving |
| `balldevice_<name>_ejecting_ball` | Eject commanded |
| `balldevice_<name>_ball_eject_success` | Ball arrived |

## Flippers & auto-fire

| Event | When |
|---|---|
| `flipper_<name>_activate` / `_deactivate` | Flipper pressed / released |
| `auto_fire_coil_<name>_activate` / `_deactivate` | Coil fired / released |

## Diverters

| Event | When |
|---|---|
| `diverter_<name>_activating` | Activating |
| `diverter_<name>_deactivating` | Deactivating |

## Drop targets

| Event | When |
|---|---|
| `drop_target_<name>_down` | Went down |
| `drop_target_<name>_up` | Came up |

## Magnets

| Event | When |
|---|---|
| `magnet_<name>_grabbing_ball` / `_grabbed_ball` | Grab started / complete |
| `magnet_<name>_releasing_ball` / `_released_ball` | Release started / complete |
| `magnet_<name>_flinging_ball` / `_flinged_ball` | Fling started / complete |

## Ball search

| Event | When |
|---|---|
| `ball_search_started` | Search began |
| `ball_search_stopped` | Search ended |
| `flipper_cradle` / `flipper_release` | Flipper held / released (search suspension) |

## Shows

| Event | When |
|---|---|
| `<showname>_<key>_unblock_queue` | A queue-blocking show finished |

Plus anything listed in a show's `EventsWhenCompleted`.

---

## Naming quirks worth knowing

A few prefixes don't match their placeholder names, which trips people up:

| Device | Event prefix | Placeholder path |
|---|---|---|
| Ball device | `balldevice_<name>_…` | `device.ball_devices.<name>.…` |
| Sequence shot | `<name>_hit`, but `sequence_shot_<name>_timeout` | — |
| Auto-fire | `auto_fire_coil_<name>_…` | — |
| Tilt | `tilt_mode_<modename>_clear` | — |

When an event doesn't fire, check the prefix before anything else. Turning on the debug log at
level Debug lists every event dispatched along with whether it had listeners — see
[Debugging](debugging.md).
