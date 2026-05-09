from .imports import Image, TypedDict


class ExcluderItem(TypedDict):
	name: str
	pics: dict[str, str]

IMG = Image.Image
FINDER = dict[str, tuple[str, str | None]]
EXCLUDER = dict[str, dict[str, ExcluderItem]]
