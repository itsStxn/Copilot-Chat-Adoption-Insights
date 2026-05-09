from .imports import io, Locator, dt, plt, df2img, base64
from .str_actions import str_normalize
from .browser import goto
from .logs import log
from vars.exports import (
	IMAGE_CROP_THRESHOLD,
	DF2IMG_OPTIONS,
	POWERBI_DOM,
	ACCOUNTS,
	COLTYPES,
	ACTUAL,
	Image,
	PAGES,
	MEDIA,
	COLS,
	IMG,
	pd,
)
from .page import (
	click_and_wait,
	search_option,
	powerbi_url,
	wait_for,
)


#* ---------------------------------------------------------------------------
#* Exceptions
#* ---------------------------------------------------------------------------

class ImageWidthMismatchError(ValueError):
	"""Raised when two images that must share a width do not."""

	def __init__(self, w1: int, w2: int) -> None:
		super().__init__(f"Images must have the same width to stack; got {w1}px and {w2}px.")


class AccountsDataError(RuntimeError):
	"""Raised when ACCOUNTS['data'] has not been initialised before use."""

	def __init__(self) -> None:
		super().__init__("ACCOUNTS['data'] is not initialised — run try_update_accounts_data() first.")


class NoFilteredAccountsError(ValueError):
	"""Raised when a filtered-accounts operation is attempted with no active filter."""


#* ---------------------------------------------------------------------------
#* Image helpers
#* ---------------------------------------------------------------------------

def img_to_b64(pic: IMG | str) -> str:
	"""
	Encode an image as a base64 PNG string.

	Args:
		pic: An IMG object, or a file path whose image will be opened,
				converted to RGBA, and cropped before encoding.

	Returns:
		Base64-encoded PNG string.
	"""
	if isinstance(pic, str):
		pic = crop_pic(Image.open(pic).convert("RGBA"))

	buf = io.BytesIO()
	pic.save(buf, format="PNG")
	buf.seek(0)
	return base64.b64encode(buf.read()).decode("utf-8")


def imgs_to_b64(pics: list[IMG]) -> list[str]:
	"""Return a list of base64-encoded PNG strings from a list of IMG objects."""
	return [img_to_b64(pic) for pic in pics]


def crop_pic(image: IMG, path: str | None = None, target_width: int = 850, increase_threshold: int = 0) -> IMG:
	"""
	Remove near-white pixels, crop to content, and resize to a target width.

	Pixels whose R, G, and B channels all exceed the crop threshold are made
	transparent. The image is then cropped to its non-transparent bounding box
	and scaled to target_width while preserving the aspect ratio.

	Args:
		image:              Input image (must be convertible to RGBA).
		path:               If provided, the result is saved to this path.
		target_width:       Output width in pixels. Defaults to 850.
		increase_threshold: Added to IMAGE_CROP_THRESHOLD to raise the cutoff.
									Defaults to 0.

	Returns:
		Processed IMG object.
	"""
	threshold = IMAGE_CROP_THRESHOLD + increase_threshold

	#* Replace near-white pixels with transparent ones so getbbox() can find content
	cleaned = [
		(255, 255, 255, 0) if all(channel > threshold for channel in pixel[:3]) else pixel
		for pixel in image.getdata()
	]
	image.putdata(cleaned)

	bbox = image.getbbox()
	if bbox:
		cropped = image.crop(bbox)
		if cropped.mode != "RGBA":
			cropped = cropped.convert("RGBA")

		#* Flatten onto a white background to remove the transparency channel
		result = Image.new("RGB", cropped.size, (255, 255, 255))
		result.paste(cropped, mask=cropped.split()[3])
	else:
		#* Nothing to crop — fall back to the original as RGB
		result = image.convert("RGB")

	if path:
		result.save(path)

	scale = target_width / result.size[0]
	target_height = int(result.size[1] * scale)
	return result.resize((target_width, target_height))


def stack_images_vertically(img1: IMG, img2: IMG, gap: int = 25) -> IMG:
	"""
	Stack two same-width images vertically with a white gap between them.

	Args:
		img1: Top image.
		img2: Bottom image.
		gap:  Vertical gap in pixels. Defaults to 25.

	Returns:
		New RGB image with img1 above img2.

	Raises:
		ImageWidthMismatchError: If img1 and img2 have different widths.
	"""
	img1 = img1.convert("RGB")
	img2 = img2.convert("RGB")

	if img1.width != img2.width:
		raise ImageWidthMismatchError(img1.width, img2.width)

	total_height = img1.height + gap + img2.height
	canvas = Image.new("RGB", (img1.width, total_height), color=(255, 255, 255))
	canvas.paste(img1, (0, 0))
	canvas.paste(img2, (0, img1.height + gap))

	return canvas


#* ---------------------------------------------------------------------------
#* Screenshot capture
#* ---------------------------------------------------------------------------

def capture(elements: list[Locator], save: bool = False) -> list[IMG]:
	"""
	Screenshot, crop, and optionally save each element in elements.

	Args:
		elements: Playwright Locator objects to capture.
		save:     Write each screenshot to a PNG file on disk. Defaults to False.

	Returns:
		List of cropped IMG objects in the same order as elements.
	"""
	log("Capturing screenshots...", left_nl=1)
	screenshots: list[IMG] = []

	for element in elements:
		element.wait_for(state="attached", timeout=5000)
		element.scroll_into_view_if_needed()
		element.wait_for(state="visible", timeout=5000)

		raw = element.screenshot(type="png", scale="device")
		screenshot = Image.open(io.BytesIO(raw)).convert("RGBA")
		cropped = crop_pic(screenshot)
		screenshots.append(cropped)

		if save:
			#* Use the sub-second timestamp fragment as a lightweight unique ID
			unique_id = hex(int(dt.now().strftime("%f")))[2:][-6:]
			screenshot.save(f"{unique_id}.png")
			log(f"Screenshot saved as {unique_id}.png")

	log("Screenshots captured.")
	return screenshots


#* ---------------------------------------------------------------------------
#* Power BI capture
#* ---------------------------------------------------------------------------

def capture_powerbi_elements(team: pd.Series, save: bool = False) -> tuple[list[IMG], str]:
	"""
	Filter the Power BI dashboard to a specific team and capture the relevant visuals.

	Reuses the cached "we mau" screenshot if it has already been captured this
	session. Applies a filtered-accounts view when a filter or custom TAM is
	active; otherwise captures the raw snapshot panel.

	Args:
		team: Series row with at least "Role" and "ID" fields.
		save: Forward the save flag to capture(). Defaults to False.

	Returns:
		Tuple of (screenshots, dashboard_url_after_filtering).
	"""
	log("Taking screenshots of the Power BI dashboard...", left_nl=1)

	powerbi = PAGES["powerbi"]
	powerbi.bring_to_front()

	#* Cache the global MAU visual — it does not change between teams
	mau_key = "we mau"
	if mau_key not in MEDIA:
		MEDIA[mau_key] = capture(wait_for(powerbi, POWERBI_DOM, at=[mau_key]), save)[0]

	#* Open the role dropdown and select this team's ID
	dropdown_btn = wait_for(
		powerbi,
		at=["dropdown button"],
		dom_attr=POWERBI_DOM,
		dynamic={"dropdown button": team["Role"]},
	)[0]
	click_and_wait(dropdown_btn, powerbi)

	dropdown = wait_for(
		powerbi,
		at=["dropdown"],
		dom_attr=POWERBI_DOM,
		dynamic={"dropdown": team["Role"]},
	)[0]

	#* Deselect everything first so only the target ID is active
	select_all = search_option(
		powerbi, dropdown, at="option", go_down=False, text_filter="Select all"
	)
	already_selected = select_all.get_attribute("aria-selected") == "true"
	click_and_wait(select_all, powerbi, clicks=1 if already_selected else 2)

	click_and_wait(
		search_option(powerbi, dropdown, at="option", text_filter=team["ID"]),
		powerbi,
	)

	url = powerbi_url()
	goto(powerbi, url)

	#* Use the filtered view when an account filter or custom TAM is in play
	needs_filter = ACCOUNTS["filter"] or team["ID"] in ACTUAL["tam"]["ID"].values
	screenshots = (
		filtered_accounts(team)
		if needs_filter
		else capture(wait_for(powerbi, POWERBI_DOM, at=["snapshot"]), save)
	)

	log("Screenshots taken.")
	return screenshots, url


#* ---------------------------------------------------------------------------
#* Progress plot
#* ---------------------------------------------------------------------------

def plot_ae_progress(team: pd.Series, show_datapoint_num: bool = False) -> list[IMG]:
	"""
	Plot Copilot Chat MAU and Incremental MAU over time for the given team.

	Filters account rows by team ID and (if set) the active account filter,
	then aggregates by date before plotting on a dual-axis line chart.

	Args:
		team:               Series row with "Role", "Names", and "ID".
		show_datapoint_num: Annotate each data point with its value. Defaults to False.

	Returns:
		Single-element list containing a cropped IMG of the generated chart.

	Raises:
		AccountsDataError: If ACCOUNTS["data"] has not been initialised.
	"""
	if ACCOUNTS["data"] is None:
		raise AccountsDataError()

	df: pd.DataFrame = ACCOUNTS["data"].copy()
	account_filter = (
		[s.lower().strip() for s in ACCOUNTS["filter"]] if ACCOUNTS["filter"] else None
	)

	role = team["Role"]
	names = team["Names"].replace(",", ", ")
	log(f"Plotting account progress for {role}/s: {names}...", left_nl=1)

	#* Narrow to this team's rows
	progress = df[df[role] == team["ID"]].copy()

	if account_filter:
		progress[COLS["account"]] = progress[COLS["account"]].apply(
			lambda x: str_normalize(str(x).strip())
		)
		progress = progress[
			progress[COLS["account"]].str.lower().str.strip().isin(account_filter)
		]

	progress[COLS["date"]]        = pd.to_datetime(progress[COLS["date"]])
	progress[COLS["adoption"]]    = pd.to_numeric(progress[COLS["adoption"]])
	progress[COLS["incremental"]] = pd.to_numeric(progress[COLS["incremental"]])

	progress = progress.groupby(COLS["date"]).agg(
		{COLS["adoption"]: "sum", COLS["incremental"]: "sum"}
	).reset_index()

	fig, ax1 = plt.subplots(figsize=(12, 6))
	x = progress[COLS["date"]]

	ax1.set_xlabel(COLS["date"])
	ax1.plot(x, progress[COLS["adoption"]], marker="o",
				label="Copilot Chat MAU (Paid + UnPaid)", color="tab:blue", linewidth=2)
	ax1.tick_params(axis="y", labelcolor="tab:blue")
	ax1.set_xticks(progress[COLS["date"]])
	ax1.set_xticklabels(progress[COLS["date"]].dt.strftime("%d-%B-%Y"), rotation=45)

	if show_datapoint_num:
		for i, value in enumerate(progress[COLS["adoption"]]):
			ax1.text(x.iloc[i], value, str(value), color="black", fontsize=9, ha="center", va="bottom")

	ax2 = ax1.twinx()
	ax2.plot(x, progress[COLS["incremental"]], marker="o",
				label=COLS["incremental"], color="tab:orange", linewidth=2)
	ax2.tick_params(axis="y", labelcolor="tab:orange")

	if show_datapoint_num:
		for i, value in enumerate(progress[COLS["incremental"]]):
			ax2.text(x.iloc[i], value, str(value), color="black", fontsize=9, ha="center", va="bottom")

	#* Merge legends from both axes into one
	handles = ax1.get_legend_handles_labels()[0] + ax2.get_legend_handles_labels()[0]
	labels  = ax1.get_legend_handles_labels()[1] + ax2.get_legend_handles_labels()[1]
	ax1.legend(handles, labels, loc="upper center", bbox_to_anchor=(0.5, 1.15), ncol=2)
	plt.tight_layout()

	buf = io.BytesIO()
	fig.savefig(buf, format="png")
	buf.seek(0)
	plt.close(fig)

	log("Account progress plot created.")
	return [crop_pic(Image.open(buf).convert("RGBA"))]


#* ---------------------------------------------------------------------------
#* Filtered accounts view
#* ---------------------------------------------------------------------------

def filtered_accounts_overview(team: pd.Series) -> tuple[pd.DataFrame, IMG]:
	"""
	Render the most-recent snapshot for the team's accounts as a table image.

	Applies the active account filter (ACCOUNTS["filter"]) when set.

	Args:
		team: Series row with at least an "ID" field.

	Returns:
		Tuple of (filtered_dataframe, cropped_table_image).

	Raises:
		AccountsDataError: If ACCOUNTS["data"] has not been initialised.
	"""
	if ACCOUNTS["data"] is None:
		raise AccountsDataError()

	df = ACCOUNTS["data"].copy()
	df[COLS["date"]] = pd.to_datetime(df[COLS["date"]])

	ae_id     = team["ID"]
	last_date = df[COLS["date"]].max()

	#* Restrict to this AE's rows for the most recent date only
	ae_df = df.loc[(df["AE"] == ae_id) & (df[COLS["date"]] == last_date)].copy()

	#* Determine which columns should be summed vs averaged/percent
	excluded_from_sum = set(COLTYPES["text"] + COLTYPES["avg"] + [COLS["date"]])
	sum_cols = [col for col in df.columns if col not in excluded_from_sum]

	ae_df.loc[:, sum_cols] = ae_df.loc[:, sum_cols].astype(int)
	COLTYPES["sum"] = sum_cols

	DF2IMG_OPTIONS["df"] = ae_df.drop(columns=[COLS["date"], "AE"])

	if ACCOUNTS["filter"]:
		account_filter = [str_normalize(s.lower().strip()) for s in ACCOUNTS["filter"]]
		filtered = DF2IMG_OPTIONS["df"].copy()
		filtered[COLS["account"]] = filtered[COLS["account"]].apply(
			lambda x: str_normalize(str(x).strip())
		)
		DF2IMG_OPTIONS["df"] = filtered[
			filtered[COLS["account"]].str.lower().str.strip().isin(account_filter)
		]

	fig = df2img.plot_dataframe(**DF2IMG_OPTIONS)
	img = Image.open(io.BytesIO(fig.to_image(format="png", width=1000, height=500, scale=2)))

	return DF2IMG_OPTIONS["df"], crop_pic(img)


def filtered_accounts(team: pd.Series) -> list[IMG]:
	"""
	Build a stacked image of the accounts overview and a totals summary table.

	Computes column-level averages and sums from the filtered DataFrame,
	formats percentage columns, and stacks the overview image above the totals.

	Args:
		team: Series row passed through to filtered_accounts_overview.

	Returns:
		Single-element list with the stacked IMG.
	"""
	log("Generating filtered accounts view...", left_nl=1)

	df, overview_pic = filtered_accounts_overview(team)
	totals_df = df.copy()

	avg_cols  = COLTYPES["avg"]
	sum_cols  = COLTYPES["sum"]
	pct_cols  = set(COLTYPES["%"])

	#* Convert percentage strings and plain floats to numeric for aggregation
	for col in avg_cols:
		if col in pct_cols:
			totals_df[col] = totals_df[col].str.rstrip("%").astype(float) / 100
		else:
			totals_df[col] = totals_df[col].astype(float) / 100

	#* Build the single-row totals DataFrame from aggregated values
	totals = pd.DataFrame({
		**{col: [totals_df[col].mean()] for col in avg_cols},
		**{col: [totals_df[col].sum()]  for col in sum_cols},
	})

	#* Re-format percentage columns as readable strings
	for col in pct_cols:
		totals[col] = totals[col].apply(lambda x: f"{x * 100:.1f}%")

	DF2IMG_OPTIONS["df"] = totals
	fig = df2img.plot_dataframe(**DF2IMG_OPTIONS)
	summary_img = Image.open(io.BytesIO(fig.to_image(format="png", width=1000, height=500, scale=2)))

	return [stack_images_vertically(overview_pic, crop_pic(summary_img))]
