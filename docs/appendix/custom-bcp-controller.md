# A local FlexDMD controller instead of BCP

The [slide & widget player](../players/slide-and-widget-player.md) is built to drive a separate
media controller process — normally the Godot Media Controller — over BCP, a local socket
protocol. If you'd rather write FlexDMD sequences directly in your table script and drive them
from the same `.Slide` / `.Widget` show config, you can — but it needs one small addition to your
table script, because GLF has no built-in FlexDMD output path.

This isn't a documented GLF feature. It's a pattern that falls directly out of how `bcpController`
is used internally, and it works because that variable is never more than an object GLF calls
known method names on.

> This page describes a pattern, not a finished library. The interface table below is complete
> against the current GLF source, but the stacking, expiry and cleanup behaviour in the example
> implementation is a design worth adopting, not something GLF provides for you — you're
> reimplementing a slice of what a real BCP media controller does, entirely in VBScript.

---

## Why this is necessary

Every place GLF plays a slide or widget — `GlfSlidePlayer.Play`, `GlfWidgetPlayer.Play`, and the
two slide/widget calls inside a running [show](../concepts/lights-and-shows.md) — has the same
shape:

```vbscript
If useBcp = True Then
    bcpController.PlaySlide slide, mode, event, action, expire, priority, kwargs
End If
```

There is no other branch, no event fired for "a slide was played," and no callback registry.
If `useBcp` is `False`, or `bcpController` was never assigned, these calls do nothing at all. There
is no config-only way to make a slide or widget player entry do something local — the entire
mechanism assumes an external process on the other end of a socket.

What makes the pattern below possible is that `bcpController` is just a variable holding an
object. GLF only ever calls a fixed set of method names on it — it doesn't check what class the
object is. Give it an object of your own with the same method names, and GLF can't tell the
difference.

## The interface GLF expects

This is the part easiest to get wrong, because **the calls aren't confined to slide and widget
code.** Once `useBcp = True`, ordinary game events touch `bcpController` too — an event player
firing any event at all routes through `Glf_BcpSendEvent`, which calls `bcpController.Send`
directly. A substitute object missing that method doesn't fail quietly; it throws the first time
any event player entry fires, which in practice means immediately.

The full set of methods `bcpController` needs, gathered from every call site including the ones
inside `bcp/bcp.vbs` itself:

| Method | Called from | Fires on |
|---|---|---|
| `PlaySlide(slide, context, calling_context, action, expire, priority, kwargs)` | Slide player, shows | A slide event |
| `PlayWidget(widget, context, calling_context, priority, expire)` | Widget player, shows | A widget event |
| `Send(message)` | `Glf_BcpSendEvent`, `Glf_BcpAddPlayer`, and the built-in player-added/next-player/ball-end listener | **Any** event player firing, any event; a player joining; every ball transition |
| `GetMessages()` | `Glf_BcpUpdate`, polled from the main game timer every 300 ms | Continuously, whether or not anything is happening |
| `Reset()` | `Glf_BcpUpdate`, on a `"hello"` message from `GetMessages()` | Only if your `GetMessages()` ever returns one |
| `ModeStart(name, priority)` | Mode start | Every mode starting |
| `ModeStop(name)` | Mode stop | Every mode stopping |
| `SendPlayerVariable(name, value, prevValue)` | Player variable changes, new-player and ball-start bookkeeping | Frequently — most scoring writes a player variable |
| `SendMachineVariable(name, value, prevValue)` | Machine variable changes | Machine variable writes |
| `ModeList()` | End of every `Glf_BcpUpdate` poll | Every 300 ms |
| `Disconnect()` | Table exit | Once |
| `RemoveSlide(slide)` | Not called by GLF's current internals | — |
| `PlaySound(sound, context, calling_context, priority)` | Not called by GLF's current internals | — |
| `StopSound(sound)` | Not called by GLF's current internals | — |
| `SlidesClear(context)` | Not called by GLF's current internals | — |

Everything above the line is live and will be called during ordinary play, not just when a slide
or widget fires — `Send`, `GetMessages`, `SendPlayerVariable`, `ModeStart`/`ModeStop`, and
`ModeList` all run continuously once `useBcp = True`. The last four are exposed on the real
controller but have no current caller anywhere in GLF; implement them as no-ops so a future GLF
version that does start calling them doesn't break your table, but nothing today depends on them.

`context` arrives with the mode's `mode_` prefix already stripped — GLF does that stripping itself
before calling `PlaySlide`/`PlayWidget`, so your version doesn't need to. `kwargs`, when present,
is a `Scripting.Dictionary` — the same tokens you'd set with `.Tokens()` on a slide, widget, or
show step.

`GetMessages()` is worth understanding specifically: `Glf_BcpUpdate` checks `IsEmpty(messages)`
first and exits immediately if so. Returning `Empty` is a complete, safe implementation — you never
need to construct a real message object, and `Reset()` will simply never be called, since nothing
ever produces the `"hello"` message that would trigger it.

## Building the substitute

The minimum viable version implements every method above — the live ones with real behaviour, the
rest as no-ops:

```vbscript
Class LocalFlexDmdController

    Public Sub PlaySlide(slide, context, calling_context, action, expire, priority, kwargs)
        If action = "remove" Then
            FlexDmd_ClearSlide slide
        Else
            FlexDmd_ShowSlide slide, kwargs
        End If
    End Sub

    Public Sub PlayWidget(widget, context, calling_context, priority, expire)
        FlexDmd_ShowWidget widget, expire
    End Sub

    Public Sub Send(message) : End Sub
    Public Function GetMessages() : GetMessages = Empty : End Function
    Public Sub Reset() : End Sub
    Public Sub ModeStart(name, priority) : End Sub
    Public Sub ModeStop(name) : End Sub
    Public Sub SendPlayerVariable(name, value, prevValue) : End Sub
    Public Sub SendMachineVariable(name, value, prevValue) : End Sub
    Public Sub ModeList() : End Sub
    Public Sub Disconnect() : End Sub

    ' Exposed on the real controller, unused by GLF today — kept for forward compatibility
    Public Sub RemoveSlide(slide) : End Sub
    Public Sub PlaySound(sound, context, calling_context, priority) : End Sub
    Public Sub StopSound(sound) : End Sub
    Public Sub SlidesClear(context) : End Sub

End Class
```

`FlexDmd_ShowSlide` and `FlexDmd_ShowWidget` are two `Select Case` blocks you write, mapping slide
and widget names to your own FlexDMD subs:

```vbscript
Sub FlexDmd_ShowSlide(slide, kwargs)
    Select Case slide
        Case "base"
            FlexDmdShowBaseSlide
        Case "jackpot_award"
            FlexDmdShowJackpotAward kwargs
        Case "high_score"
            FlexDmdShowHighScoreEntry kwargs
    End Select
End Sub

Sub FlexDmd_ShowWidget(widget, expireSeconds)
    Select Case widget
        Case "ball_save"
            FlexDmdFlashBallSaved expireSeconds
    End Select
End Sub
```

Each `Case` is a slide or widget name you'll also reference from `.Slide = "..."` /
`.Widget = "..."` in your mode config — same names, same place they'd go for a real BCP setup.
Everything past the `Select Case` is ordinary FlexDMD code with no further GLF involvement.

## Going further

The minimum version above is enough to make slides and widgets appear, one at a time, with no
memory of what came before. It doesn't handle two things wanting the display at once, an `Expire`
timeout, or a mode's slides disappearing when the mode stops — which matters, because GLF's
`priority`, `context`, and `Expire` parameters are clearly designed around exactly that, the same
way the [light stack](../concepts/lights-and-shows.md#the-light-stack) and the
[segment display stack](../players/segment-display-player.md#the-display-stack) behave. GLF just
doesn't implement the equivalent for slides and widgets — that's left entirely to whatever's behind
`bcpController`.

Building that is a real, non-trivial piece of code — priority stacking, tie-breaking, timed
removal, and per-mode cleanup all interact. Below is a complete reference implementation covering
all of it, adapted from a working table.

## Reference implementation

### The stack

Each active slide is one entry: a priority, the mode (`context`) that played it, and a sequence
number used to break ties — when two slides share a priority, the more recently played one wins,
which matches how the light stack resolves ties by recency.

```vbscript
Class GlfFlexDmdSlide

    Private m_priority, m_context, m_seq, m_kwargs

    Public Property Get Priority() : Priority = m_priority : End Property
    Public Property Let Priority(input) : m_priority = input : End Property

    Public Property Get Context() : Context = m_context : End Property
    Public Property Let Context(input) : m_context = input : End Property

    Public Property Get Seq() : Seq = m_seq : End Property
    Public Property Let Seq(input) : m_seq = input : End Property

    Public default Function Init()
        m_priority = 0
        m_context = ""
        m_seq = 0
        m_kwargs = Null
        Set Init = Me
    End Function

    ' kwargs is whatever the event carried - a Scripting.Dictionary, or
    ' Null from a show step. Property Let cannot take an object, hence
    ' the pair of methods.
    Public Sub SetKwargs(input)
        If IsObject(input) Then
            Set m_kwargs = input
        Else
            m_kwargs = Null
        End If
    End Sub

    Public Property Get HasKwargs()
        HasKwargs = IsObject(m_kwargs)
    End Property

    Public Function Kwargs()
        If IsObject(m_kwargs) Then
            Set Kwargs = m_kwargs
        Else
            Kwargs = Null
        End If
    End Function

End Class
```

### The controller

```vbscript
Class GlfFlexDmdBcpController

    Private m_slides     ' slide name -> GlfFlexDmdSlide
    Private m_current    ' slide name currently rendered ("" = none)
    Private m_seq        ' monotonic counter, breaks priority ties
    Private m_connected

    Public default Function Init()
        Set m_slides = CreateObject("Scripting.Dictionary")
        m_current = ""
        m_seq = 0
        m_connected = True
        Set Init = Me
    End Function

    ' False once Disconnect has run - lets FlexBcp_Attach (below) tell a
    ' live controller from a spent one.
    Public Property Get Connected() : Connected = m_connected : End Property


    '--- Slides --------------------------------------------------------

    Public Sub PlaySlide(slide, context, calling_context, action, expire, priority, kwargs)
        If m_connected = False Then Exit Sub
        If FlexBcp_Str(slide) = "" Then Exit Sub

        If LCase(FlexBcp_Str(action)) = "remove" Then
            RemoveSlide slide
            Exit Sub
        End If

        Dim entry
        If m_slides.Exists(slide) Then
            Set entry = m_slides(slide)
        Else
            Set entry = (new GlfFlexDmdSlide)()
            m_slides.Add slide, entry
        End If

        m_seq = m_seq + 1
        entry.Priority = FlexBcp_Num(priority)
        entry.Context = FlexBcp_StripMode(context)
        entry.Seq = m_seq
        entry.SetKwargs kwargs

        RemoveDelay FlexBcp_ExpiryKey(slide)
        If FlexBcp_Num(expire) > 0 Then
            SetDelay FlexBcp_ExpiryKey(slide), "Glf_FlexDmdSlideExpired", slide, _
                     FlexBcp_Num(expire) * 1000
        End If

        ' Force a re-render if this slide is the top one, even when it
        ' already was - replaying a slide should restart its animation.
        Render slide
    End Sub

    Public Sub RemoveSlide(slide)
        If m_connected = False Then Exit Sub
        If m_slides.Exists(slide) Then
            m_slides.Remove slide
            RemoveDelay FlexBcp_ExpiryKey(slide)
        End If
        Render ""
    End Sub

    ' Drop every slide a given mode played. GLF has no internal caller for
    ' this method, but ModeStop below calls it directly, so a mode's
    ' slides clean themselves up without needing an event player entry.
    Public Sub SlidesClear(context)
        If m_connected = False Then Exit Sub
        If m_slides.Count = 0 Then Exit Sub

        Dim ctx : ctx = FlexBcp_StripMode(context)
        Dim doomed() : ReDim doomed(m_slides.Count)
        Dim n : n = -1
        Dim slideName, entry
        For Each slideName In m_slides.Keys()
            Set entry = m_slides(slideName)
            If entry.Context = ctx Then
                n = n + 1
                doomed(n) = slideName
            End If
        Next
        If n = -1 Then Exit Sub

        Dim i
        For i = 0 To n
            m_slides.Remove doomed(i)
            RemoveDelay FlexBcp_ExpiryKey(doomed(i))
        Next
        Render ""
    End Sub


    '--- Widgets -------------------------------------------------------

    Public Sub PlayWidget(widget, context, calling_context, priority, expire)
        If m_connected = False Then Exit Sub
        If FlexBcp_Str(widget) = "" Then Exit Sub

        Dim secs : secs = FlexBcp_Num(expire)
        If secs <= 0 Then secs = FlexDmdDefaultWidgetExpire   ' see "Frame timing" below

        FlexDmd_ShowWidget widget, secs
    End Sub


    '--- Modes -----------------------------------------------------------
    ' ModeStop fires for every mode, every time it stops, whether or not
    ' that mode ever played a slide - a stronger guarantee than relying
    ' on SlidesClear ever being called directly, since nothing in GLF's
    ' current internals calls it.

    Public Sub ModeStart(name, priority) : End Sub
    Public Sub ModeStop(name) : SlidesClear name : End Sub
    Public Sub ModeList() : End Sub


    '--- Sound -------------------------------------------------------------
    ' GLF's own SoundPlayer handles audio locally and never routes it
    ' through bcpController, so there's nothing to do here.

    Public Sub PlaySound(sound, context, calling_context, priority) : End Sub
    Public Sub StopSound(sound) : End Sub


    '--- Raw BCP surface -----------------------------------------------
    ' Once useBcp is True, GLF pushes every dispatched event, player add,
    ' ball start/end and variable change at these. They're wire-protocol
    ' messages for a process that doesn't exist here, so they're dropped -
    ' but they must exist, or those call sites error out.

    Public Sub Send(commandMessage) : End Sub

    ' Glf_BcpUpdate polls this every 300ms and returns early on Empty.
    Public Function GetMessages()
        GetMessages = Empty
    End Function

    Public Sub Reset() : End Sub
    Public Sub SendPlayerVariable(name, value, prevValue) : End Sub
    Public Sub SendMachineVariable(name, value, prevValue) : End Sub

    Public Sub Disconnect()
        If m_connected Then
            m_connected = False
            useBcp = False
            m_slides.RemoveAll
            m_current = ""
        End If
    End Sub


    '--- Internals -----------------------------------------------------

    ' Render the top of the stack. forceSlide re-renders even when that
    ' slide is already the one showing; pass "" for "only if it changed".
    Private Sub Render(forceSlide)
        Dim topSlideName : topSlideName = TopSlide()

        If topSlideName = "" Then
            ' Nothing left in the stack. Forget the current slide so a
            ' later replay of it renders again, but leave whatever's
            ' already on the DMD alone - see "Leaving the last scene up"
            ' below.
            m_current = ""
            Exit Sub
        End If

        If topSlideName <> m_current Or topSlideName = forceSlide Then
            Dim entry : Set entry = m_slides(topSlideName)
            Dim kw
            If entry.HasKwargs Then
                Set kw = entry.Kwargs()
            Else
                kw = Null
            End If

            FlexDmd_ShowSlide topSlideName, kw
            m_current = topSlideName
        End If
    End Sub

    ' Highest priority wins; most recently played breaks a tie.
    Private Function TopSlide()
        Dim bestName : bestName = ""
        Dim bestPri, bestSeq
        Dim slideName, entry
        For Each slideName In m_slides.Keys()
            Set entry = m_slides(slideName)
            If bestName = "" Or entry.Priority > bestPri Or _
               (entry.Priority = bestPri And entry.Seq > bestSeq) Then
                bestName = slideName
                bestPri = entry.Priority
                bestSeq = entry.Seq
            End If
        Next
        TopSlide = bestName
    End Function

End Class


' Fired by SetDelay when a slide's Expire elapses.
Sub Glf_FlexDmdSlideExpired(args)
    If IsObject(glfFlexBcp) Then glfFlexBcp.RemoveSlide args
End Sub
```

`FlexDmd_ShowSlide` and `FlexDmd_ShowWidget` are the same two `Select Case` blocks from the minimum
version above — nothing about the mapping layer changes; only the controller managing *when* they
get called has grown.

### Design notes

**Leaving the last scene up.** When the stack empties — every slide removed or expired — `Render`
deliberately does nothing further, rather than blanking the display. In practice this means the
scoreboard (or whatever was showing) stays visible between balls, when the base mode has stopped
and nothing new has played a slide yet. Blanking on empty is a one-line change in `Render` if you'd
rather the display go dark.

**`SetDelay` for expiry.** `Expire` is a duration in seconds; `SetDelay` is GLF's own one-shot timer
utility (see [the event system](../concepts/events.md#delays)). Each slide's expiry timer is keyed
by its own name, so replaying a slide before its previous expiry fires correctly resets the clock —
`RemoveDelay` followed by a fresh `SetDelay` on every `PlaySlide` call handles that.

**Why `Expire` isn't handled the same way for widgets.** The reference class above hands widgets
straight to `FlexDmd_ShowWidget` with no stack entry of their own. That's a simplification, not an
oversight: a widget is meant to be a transient overlay on the current slide, not a competitor for
priority — if your FlexDMD scenes are structured so a widget can render over any slide, this is
enough. If you want widgets themselves priority-stacked (two widgets fighting for the same
overlay), the same `GlfFlexDmdSlide` pattern applies equally to them.

**Frame timing.** `Expire` arrives in seconds, but FlexDMD helpers like `DMDBigText` typically count
frames of your render timer, not seconds. Keep a constant tied to your actual timer interval and
convert at the boundary, rather than hard-coding frame counts per call:

```vbscript
Const FlexDmdFrameMs = 17   ' must match your DMDTimer's Interval
Const FlexDmdDefaultWidgetExpire = 1.3

Function FlexDmd_Frames(seconds)
    FlexDmd_Frames = Int(seconds * 1000 / FlexDmdFrameMs)
End Function
```

If this constant drifts out of sync with the real timer object's interval, every duration silently
runs long or short with nothing to signal the mismatch — worth a comment at both ends pointing at
each other.

### Helpers

```vbscript
Function FlexBcp_Num(value)
    If IsObject(value) Then
        FlexBcp_Num = 0
    ElseIf IsEmpty(value) Or IsNull(value) Then
        FlexBcp_Num = 0
    ElseIf IsNumeric(value) Then
        FlexBcp_Num = CDbl(value)
    Else
        FlexBcp_Num = 0
    End If
End Function

Function FlexBcp_Str(value)
    If IsObject(value) Then
        FlexBcp_Str = ""
    ElseIf IsEmpty(value) Or IsNull(value) Then
        FlexBcp_Str = ""
    Else
        FlexBcp_Str = CStr(value)
    End If
End Function

' The slide player passes the mode name unprefixed already; a show step
' can pass it with "mode_" still on. Normalise the way GLF's own BCP
' controller does, so ModeStop's clear matches what PlaySlide stored.
Function FlexBcp_StripMode(value)
    FlexBcp_StripMode = Replace(FlexBcp_Str(value), "mode_", "")
End Function

Function FlexBcp_ExpiryKey(slide)
    FlexBcp_ExpiryKey = "flexdmd_slide_expire_" & slide
End Function
```

`FlexBcp_Num` and `FlexBcp_Str` exist because `priority`, `expire`, `context` and similar arguments
can arrive as `Empty` from a show step that never set them — coercing blind with `CDbl`/`CStr`
throws on `Empty`, where these return a safe default instead.

### Attaching it safely

The minimum version's `Sub Table1_Init` / `Sub Table1_OptionEvent` pairing (shown above) works, but
it unconditionally overwrites `bcpController` on every call — fine if nothing else in your table
ever touches that variable, riskier if it might. A small wrapper makes attaching idempotent and
leaves someone else's controller alone if one is already there:

```vbscript
Dim glfFlexBcp : glfFlexBcp = Null   ' our controller instance, or Null

Sub FlexBcp_Attach()
    If UseFlexDMD = 0 Then Exit Sub   ' your table's own FlexDMD-enabled flag, not a GLF variable

    ' Real Godot controller wins if the table option asks for it.
    If Table1.Option("Glf Backbox Control Protocol", 0, 1, 1, 0, 0, Array("Off", "On")) = 1 Then
        Exit Sub
    End If

    If IsObject(bcpController) Then
        Dim isOurs : isOurs = False
        If IsObject(glfFlexBcp) Then
            If bcpController Is glfFlexBcp Then isOurs = True
        End If
        If Not isOurs Then Exit Sub   ' somebody else's controller - leave it

        If bcpController.Connected Then
            useBcp = True
            Exit Sub                  ' already ours and live
        End If
        ' ours, but Disconnect has run - fall through and replace it
    End If

    Set glfFlexBcp = (new GlfFlexDmdBcpController)()
    Set bcpController = glfFlexBcp
    useBcp = True
End Sub
```

Call `FlexBcp_Attach()` from both `Table1_Init` and `Table1_OptionEvent` — shown in
[Wiring it in](#wiring-it-in) below, along with why the second call site matters.

## Wiring it in

Skip GLF's normal BCP connection path — don't call `Glf_ConnectToBCPMediaController`, and leave the
*Glf Backbox Control Protocol* table option off (both would try to launch and connect to a real
Godot process). If you're using the minimum viable class from above, assign it directly:

```vbscript
Sub Table1_Init
    ConfigureGlfDevices()
    Glf_Init(Table1)

    Set bcpController = New LocalFlexDmdController
    useBcp = True
End Sub
```

If you're using the reference implementation, call `FlexBcp_Attach()` instead — it handles the same
assignment, plus the idempotency and identity checks described above:

```vbscript
Sub Table1_Init
    ConfigureGlfDevices()
    Glf_Init(Table1)
    FlexBcp_Attach()
End Sub
```

**Re-attach after every options dialog interaction, either way.** `Glf_Options` disconnects and
nulls `bcpController` on every option event whenever the *Glf Backbox Control Protocol* table
option is off — which it will be, since you're not using it. Left alone, that means your display
goes dead the first time anyone opens the table's tweak/options menu:

```vbscript
Sub Table1_OptionEvent(ByVal eventId)
    Glf_Options(eventId)
    FlexBcp_Attach()          ' or: Set bcpController = New LocalFlexDmdController : useBcp = True
End Sub
```

From here, every slide, widget, and show configured the normal way —
[show player](../players/show-player.md), [slide & widget player](../players/slide-and-widget-player.md) —
routes to your FlexDMD subs instead of a socket, with no change to how you write mode config.

```vbscript
With .SlidePlayer()
    With .EventName("mode_base_started")
        .Slide = "base"
        .Action = "play"
    End With
End With

With .WidgetPlayer()
    With .EventName("ball_save_new_ball_saving_ball")
        .Widget = "ball_save"
        .Action = "play"
        .Expire = 2
    End With
End With
```

Both fire exactly as documented — the only difference from a real Godot setup is what's sitting
behind `bcpController`.

## Migrating existing display code incrementally

If your table already drives FlexDMD from `AddPinEventListener` calls scattered through table-side
handlers, there's no need to move everything across at once. Once `bcpController` is wired in, move
one slide or widget at a time: add the `.Slide` / `.Widget` entry to the relevant mode's config,
confirm your `Select Case` block handles that name, and then delete the old listener and its
callback so the two paths don't both fire for the same trigger. The rest keep working exactly as
before until you get to them.

## Reverting to real BCP later

Nothing about this pattern touches mode config, so switching back to a real Godot Media Controller
later is just changing what gets assigned to `bcpController` — call
`Glf_ConnectToBCPMediaController` the normal way instead of constructing your own class (and turn
the *Glf Backbox Control Protocol* table option back on so `Glf_Options` doesn't disconnect it),
and every show, slide, and widget entry keeps working unmodified.

## See also

[Slide & widget player](../players/slide-and-widget-player.md) ·
[Lights & shows](../concepts/lights-and-shows.md) · [Modes](../concepts/modes.md) ·
[The event system](../concepts/events.md)
