from enum import Enum, auto
from typing import Callable

subscribers = {}

class EventType(Enum):
    ON_START_DATE_VAR_CHANGE = auto()
    ON_END_DATE_VAR_CHANGE = auto()


def subscribe(event_type: EventType, fn: Callable[[], None]) -> None:
    if event_type not in subscribers:
        subscribers[event_type] = []
    subscribers[event_type].append(fn)


def postEvent(event_type: EventType) -> None:
    if event_type not in subscribers:
        return
    for fn in subscribers[event_type]:
        fn()

