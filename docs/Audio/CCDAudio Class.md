# CCDAudio Class

The `CCDAudio` class allows to play a CD Rom using MCI.

### Constructor

| Name       | Description |
| ---------- | ----------- |
| [Constructor](#constructor) | Creates an instance of the `CCDAudio` class. |

---

### Methods

| Name       | Description |
| ---------- | ----------- |
| [Backward](#backward) | Moves to the previous track. |
| [Close](#close) | Closes the device or file and any associated resources. MCI unloads a device when all instances of the device or all files are closed.  |
| [CloseDoor](#closedoor) | Closes the CDRom door. |
| [Forward](#forward) | Moves to the next track. |
| [GetAllTracksLength](#getalltrackslength) | Returns the total length in seconds of all the tracks. |
| [GetAllTracksLengthString](#getalltrackslengthstring) | Returns the total length of all the tracks. |
| [GetCurrentPos](#getcurrentpos) | Returns the current track position in seconds. |
| [GetCurrentPosString](#getcurrentposstring) | Returns the current track position. |
| [GetCurrentTrack](#getcurrenttrack) | Returns the current track number. |
| [GetErrorString](#geterrorstring) | Retrieves a string that describes the specified MCI error code. |
| [GetLastError](#getlasterror) | Retrieves a The last MCI error code. |
| [GetTrackLength](#gettracklength) | Returns the length in seconds of the given track. |
| [GetTrackLengthString](#gettracklengthstring) | Returns the length of the given track. |
| [GetTracksCount](#gettrackscount) | Returns the count of tracks. |
| [GetTrackStartTime](#gettrackstarttime) | Returns the start time of the given track. |
| [GetTrackStartTimeString](#gettrackstarttimestring) | Returns the start time of the given track. |
| [IsMediaInserted](#ismediainserted) | Checks whether CD media is inserted. |
| [IsPaused](#ispaused) | Checks whether is in paused mode. |
| [IsPlaying](#isplaying) | Checks whether is in play mode. |
| [IsReady](#isready) | Checks if the device is ready. |
| [IsSeeking](#isseeking) | Checks whether is in seeking mode. |
| [IsStopped](#ssstopped) | Checks whether is in stopped mode. |
| [Open](#open) | Initializes the device. |
| [OpenDoor](#opendoor) | Opens the CDRom door. |
| [Pause](#pause) | Pauses playing CD Audio. |
| [Play](#play) | Starts playing CD Audio. |
| [PlayFrom](#playfrom) | Starts playing CD Audio on the given track. |
| [PlayFromTo](#playfromto) | Starts playing CD Audio from a given track to a given track. |
| [Stop](#dtop) | Stops playing CD Audio. |
| [ToEnd](#toend) | Sets the position to the end of the audio CD. |
| [ToStart](#tostart) | Sets the position to the start of the audio CD. |

---

## Constructor

Creates an instance of the `CCDAudio` class.

```
CONSTRUCTOR CCDAudio
```

#### Usage example:

```
DIM pAudio AS CCDAudio
pAudio.Open
pAudio.Play
```
---

## Backward

Moves to the previous track.

```
FUNCTION Backward () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

---

## Close

Closes the device or file and any associated resources. MCI unloads a device when all instances of the device or all files are closed. 

```
FUNCTION Close () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

---

## CloseDoor

Closes the CDRom door.

```
FUNCTION CloseDoor () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

---

## Forward

Moves to the next track.

```
FUNCTION Forward () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

---

## GetAllTracksLength

Returns the total length in seconds of all the tracks.

```
FUNCTION GetAllTracksLength () AS LONG
```
---

## GetAllTracksLengthString

Returns the total length of all the tracks as a string.

```
FUNCTION GetAllTracksLengthString () AS CWSTR
```
---

## GetCurrentPos

Returns the current track position in seconds.

```
FUNCTION GetCurrentPos () AS LONG
```
---

## GetCurrentPosString

Returns the current track position as a string.

```
FUNCTION GetCurrentPosString () AS CWSTR
```
---

## GetCurrentTrack

Returns the current track number.

```
FUNCTION GetCurrentTrack () AS LONG
```
---

## GetErrorString

Retrieves a string that describes the specified MCI error code.

```
FUNCTION GetErrorString (BYVAL dwError AS MCIERROR = 0) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *dwError* | Optional. The MCI error code. If this parameter is omitted, the last error code is used. |

---

## GetLastError

Returns the last MCI error code.

```
FUNCTION GetLastError () AS MCIERROR
```
---

## GetTrackLength

Returns the length in seconds of the given track.

```
FUNCTION GetTrackLength (BYVAL nTrack AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

---

## GetTrackLengthString

Returns the length of the given track as a string.

```
FUNCTION GetTrackLengthString (BYVAL nTrack AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

---

## GetTracksCount

Returns the count of tracks.

```
FUNCTION GetTracksCount () AS LONG
```
---

## GetTrackStartTime

Returns the start time of the given track.

```
FUNCTION GetTrackStartTime (BYVAL nTrack AS LONG) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

---

## GetTrackStartTimeString

Returns the start time of the given track as a string.

```
FUNCTION GetTrackStartTimeString (BYVAL nTrack AS LONG) AS CWSTR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

---

## IsMediaInserted

Checks whether CD media is inserted.

```
FUNCTION IsMediaInserted () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

---

## IsPaused

Checks whether is in paused mode.

```
FUNCTION IsPaused () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

---

## IsPlaying

Checks whether is in play mode.

```
FUNCTION IsPlaying () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

---

## IsReady

Checks if the device is ready.

```
FUNCTION IsReady () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

---

## IsSeeking

Checks whether is in seeking mode.

```
FUNCTION IsSeeking () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

---

## IsStopped

Checks whether is in stopped mode.

```
FUNCTION IsStopped () AS BOOLEAN
```

#### Return value

TRUE or FALSE.

---

## Open

Initializes the device.

```
FUNCTION Open () AS DWORD
```

#### Return value

Returns zero if successful or an error otherwise.

---

## OpenDoor

Opens the CDRom door.

```
FUNCTION OpenDoor () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

---

## Pause

Pauses playing CD Audio.

```
FUNCTION Pause () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

---

## Play

Starts playing CD Audio.

```
FUNCTION Play () AS MCIERROR
```

#### Return value

Returns zero if successful or an error otherwise.

---

## PlayFrom

Starts playing CD Audio on the given track.

```
FUNCTION PlayFrom (BYVAL nTrack AS LONG) AS MCIERROR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nTrack* | The track number. |

#### Return value

Returns zero if successful or an error otherwise.

---

## PlayFromTo

Starts playing CD Audio from a given track to a given track.

```
FUNCTION PlayFromTo (BYVAL nStartTrack AS LONG, BYVAL nEndTrack AS LONG) AS MCIERROR
```

| Parameter  | Description |
| ---------- | ----------- |
| *nStartTrack* | The starting track number. |
| *nEndTrack* | The ending track number. |

#### Return value

Returns zero if successful or an error otherwise.

---

## Stop

Stops playing CD Audio.

```
FUNCTION Stop () AS MCIERROR
```
---

## ToEnd

Sets the position to the end of the audio CD.

```
FUNCTION ToEnd () AS MCIERROR
```
---

## ToStart

Sets the position to the start of the audio CD.

```
FUNCTION ToStart () AS MCIERROR
```
---
