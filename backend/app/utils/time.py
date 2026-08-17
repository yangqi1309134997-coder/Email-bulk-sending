"""Time helpers shared by the API, workers, and persistence models.

The application stores timestamps as naive UTC values for compatibility with
the existing SQLite/PostgreSQL schema.  ``datetime.utcnow`` is deprecated in
modern Python, so keep the storage contract while obtaining the value through
the timezone-aware API.
"""

from datetime import datetime, timezone


def utcnow() -> datetime:
    """Return the current UTC time as a naive datetime.

    Database columns in this project historically contain naive UTC values.
    Converting at this boundary avoids mixing aware and naive datetimes while
    removing reliance on the deprecated ``datetime.utcnow`` function.
    """

    return datetime.now(timezone.utc).replace(tzinfo=None)


def from_unix_utc(timestamp: float) -> datetime:
    """Convert a Unix timestamp to the application's naive UTC format."""

    return datetime.fromtimestamp(timestamp, timezone.utc).replace(tzinfo=None)


def to_unix_utc(value: datetime) -> float:
    """Convert a stored naive/aware UTC datetime to a Unix timestamp."""

    if value.tzinfo is None:
        value = value.replace(tzinfo=timezone.utc)
    else:
        value = value.astimezone(timezone.utc)
    return value.timestamp()
