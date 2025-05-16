from .imports import io, Locator, dt, plt, df2img, base64
from .str_actions import str_normalize
from .browser import goto
from .logs import log
from vars.exports import (
    IMAGE_CROP_THRESHOLD, 
    DF2IMG_OPTIONS,
    POWERBI_DOM, 
    ACCOUNTS, 
    Image, 
    PAGES, 
    MEDIA, 
    COLS, 
    IMG, 
    pd
)
from .page import (
    click_and_wait, 
    search_option,
    powerbi_url, 
    wait_for
)


def img_to_b64(pic:IMG|str) -> str:
    """
    Converts an image to a base64-encoded PNG string.
    Args:
        pic (IMG or str): The image to convert. Can be an image object (IMG) or a file path (str).
            If a file path is provided, the image is opened, converted to RGBA, and cropped using `crop_pic`.
    Returns:
        str: The base64-encoded string representation of the image in PNG format.
    """

    if isinstance(pic, str):
        pic = crop_pic(Image.open(pic).convert("RGBA"))

    buffer = io.BytesIO()
    pic.save(buffer, format='PNG')
    buffer.seek(0)
    encoded = base64.b64encode(buffer.read()).decode('utf-8')

    return encoded

def imgs_to_b64(pics:list[IMG]) -> list[str]:
    """
    Converts a list of IMG objects to a list of their base64-encoded string representations.
    Args:
        pics (list[IMG]): A list of IMG objects to be converted.
    Returns:
        list[str]: A list of base64-encoded strings corresponding to the input images.
    """
    
    pics_b64 = []
    for pic in pics:
        pics_b64.append(img_to_b64(pic))

    return pics_b64

def crop_pic(image:IMG, path=None, target_width=850, increase_threshold=0) -> IMG:
    """
    Crops and resizes an image by removing pixels above a brightness threshold and making them transparent, then cropping to the non-transparent area and resizing to a target width.
    Args:
        image (IMG): The input image to be processed.
        path (str, optional): If provided, saves the resulting image to this file path.
        target_width (int, optional): The desired width of the output image. Defaults to 850.
        increase_threshold (int, optional): Value to increase the crop threshold by. Defaults to 0.
    Returns:
        IMG: The processed and resized image.
    Notes:
        - Pixels with all RGB values above the threshold are made transparent.
        - The image is cropped to the bounding box of non-transparent pixels.
        - The result is resized to the specified width, maintaining aspect ratio.
        - If a path is provided, the result is saved to that location.
    """

    new_data = []
    for pixel in image.getdata():
        thres = IMAGE_CROP_THRESHOLD + increase_threshold
        if pixel[0] > thres and pixel[1] > thres and pixel[2] > thres:
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(pixel)

    image.putdata(new_data)

    result = None
    bbox = image.getbbox()
    
    if bbox:
        cropped_img = image.crop(bbox)
        if cropped_img.mode != "RGBA":
            cropped_img = cropped_img.convert("RGBA")
        
        background = Image.new("RGB", cropped_img.size, (255, 255, 255))
        background.paste(cropped_img, mask=cropped_img.split()[3])
        if path: background.save(path)

        result = background
    else:
        result = image.convert("RGBA")
        
    w_percent = target_width / float(result.size[0])
    target_height = int(result.size[1] * w_percent)
    result = result.resize((target_width, target_height), Image.LANCZOS)

    return result

def capture(elements:list[Locator], save=False) -> list[IMG]:
    """
    Captures screenshots of the provided list of Locator elements.
    Args:
        elements (list[Locator]): A list of Locator objects to capture screenshots from.
        save (bool, optional): If True, saves each screenshot as a PNG file with a unique ID. Defaults to False.
    Returns:
        list[IMG]: A list of cropped screenshot images (IMG objects).
    Side Effects:
        - Logs progress and status messages.
        - Optionally saves screenshots to disk as PNG files.
    Raises:
        Any exceptions raised by Locator methods or image processing will propagate.
    """

    log("Capturing screenshots...", left_nl=1)

    screenshots = []
    for i, element in enumerate(elements):
        element.wait_for(state="attached", timeout=5000)
        element.scroll_into_view_if_needed()
        element.wait_for(state="visible", timeout=5000)

        id = hex(int(dt.now().strftime("%f")))[2:][-6:]

        pic = element.screenshot(type="png", scale="device")
        screenshot = Image.open(io.BytesIO(pic)).convert("RGBA")
        screenshots.append(crop_pic(screenshot))

        if save:
            screenshot.save(f"{id}.png")
            log(f"Screenshot saved as {id}.png")

    log("Screenshots captured")

    return screenshots

def stack_images_vertically(img1: IMG, img2: IMG, gap=25) -> IMG:
    """
    Stacks two images vertically with an optional gap between them.
    Args:
        img1 (IMG): The first image to stack (appears on top).
        img2 (IMG): The second image to stack (appears below).
        gap (int, optional): The number of pixels to place as a gap between the two images. Defaults to 25.
    Returns:
        IMG: A new image with img1 and img2 stacked vertically with the specified gap.
    Raises:
        Exception: If the widths of img1 and img2 do not match.
    Note:
        Both images are converted to RGB mode before stacking.
    """

    log("Stacking images vertically...", left_nl=1)

    img1 = img1.convert("RGB")
    img2 = img2.convert("RGB")

    if img1.width != img2.width:
        raise Exception("Images must have the same width")

    total_height = img1.height + gap + img2.height
    stacked_img = Image.new("RGB", (img1.width, total_height), color=(255, 255, 255))

    stacked_img.paste(img1, (0, 0))
    stacked_img.paste(img2, (0, img1.height + gap))

    log("Images stacked")

    return stacked_img

def capture_powerbi_elements(team:pd.Series, save=False) -> tuple[list[IMG], str]:
    """
    Captures screenshots of specific elements from a PowerBI dashboard for a given team and returns the images along with the dashboard URL.
    Args:
        team (pd.Series): A pandas Series containing team information, including "Role" and "ID" fields used for interacting with the dashboard.
        save (bool, optional): If True, saves the captured screenshots to disk. Defaults to False.
    Returns:
        tuple[list[IMG], str]: 
            - A list of captured images (screenshots) of the relevant PowerBI dashboard elements.
            - The URL of the PowerBI dashboard after filtering for the specified team.
    Raises:
        KeyError: If required keys are missing in the `team` Series.
        Exception: If any step in the interaction or capture process fails.
    Side Effects:
        - Interacts with the PowerBI dashboard UI (clicks, dropdown selections).
        - Optionally saves screenshots to disk.
        - Logs progress messages.
    """

    log("Taking screenshots of the PowerBI dashboard...", left_nl=1)

    powerbi = PAGES["powerbi"]
    powerbi.bring_to_front()

    desc = "we mau"
    if desc not in MEDIA:
        MEDIA[desc] = capture(
            wait_for(
                powerbi, 
                POWERBI_DOM, 
                at=[desc]
            ), 
            save
        )[0]

    desc = "dropdown button"
    dropdown_btn = wait_for(
        powerbi, 
        at=[desc], 
        dom_attr=POWERBI_DOM, 
        dynamic={desc: team["Role"]}
    )[0]

    click_and_wait(dropdown_btn, powerbi)

    desc = "dropdown"
    dropdown = wait_for(
        powerbi, 
        at=[desc], 
        dom_attr=POWERBI_DOM, 
        dynamic={desc: team["Role"]}
    )[0]

    desc = "option"
    select_all = search_option(
        powerbi, 
        dropdown, 
        at=desc,
        go_down=False,
        text_filter="Select all"
    )

    click_and_wait(
        select_all, 
        powerbi, 
        clicks=1 if select_all.get_attribute(
            "aria-selected"
        ) == "true" else 2
    )
    
    click_and_wait(
        search_option(
            powerbi, 
            dropdown, 
            at=desc,
            text_filter=team["ID"]
        ), 
        powerbi
    )

    url = powerbi_url()
    goto(powerbi, url)

    screenshot = filtered_accounts(team) if ACCOUNTS["filter"] else capture(
        wait_for(
            powerbi, 
            POWERBI_DOM, 
            at=["snapshot"]
        ), 
        save
    )
    
    log("Screenshots taken")

    return screenshot, url

def plot_ae_progress(team:pd.Series, show_datapoint_num=False) -> list[IMG]:
    """
    Plots the account executive (AE) progress for a given team, visualizing Copilot Chat MAU (Paid + UnPaid) 
    and Copilot Chat H2 Incremental MAU over time.
    Args:
        team (pd.Series): A pandas Series containing team information, including 'Role', 'Names', and 'ID'.
        show_datapoint_num (bool, optional): If True, displays the numerical value of each data point on the plot. Defaults to False.
    Returns:
        list[IMG]: A list containing a single cropped image (PIL Image) of the generated plot.
    """

    df: pd.DataFrame = ACCOUNTS["data"].copy()
    filter = [s.lower().strip() for s in ACCOUNTS["filter"]] if ACCOUNTS["filter"] else None

    role = team["Role"]
    names = team["Names"].replace(",", ", ")
    log(f"Plotting account progress for {role}/s: {names}...", left_nl=1)

    progress = df[df[role] == team["ID"]].copy()

    if filter:
        progress["TopParent"] = progress["TopParent"].apply(
            lambda x: str_normalize(str(x).strip())
        )
        progress = progress[progress["TopParent"].str.lower().str.strip().isin(filter)]
    
    progress['Date'] = pd.to_datetime(progress['Date'])
    progress['Copilot Chat MAU (Paid +UnPaid)'] = pd.to_numeric(progress['Copilot Chat MAU (Paid +UnPaid)'])
    progress['Copilot Chat H2 Incremental MAU'] = pd.to_numeric(progress['Copilot Chat H2 Incremental MAU'])

    progress = progress.groupby('Date').agg({
        'Copilot Chat MAU (Paid +UnPaid)': 'sum',
        'Copilot Chat H2 Incremental MAU': 'sum'
    }).reset_index()

    fig, ax1 = plt.subplots(figsize=(12, 6))
    x = progress['Date']

    ax1.set_xlabel('Date')
    ax1.plot(x, progress["Copilot Chat MAU (Paid +UnPaid)"], marker="o", label="Copilot Chat MAU (Paid + UnPaid)", color='tab:blue', linewidth=2)
    ax1.tick_params(axis='y', labelcolor='tab:blue')
    ax1.set_xticks(progress['Date'])
    ax1.set_xticklabels(progress['Date'].dt.strftime('%d-%B-%Y'), rotation=45)

    if show_datapoint_num:
        for i, value in enumerate(progress["Copilot Chat MAU (Paid +UnPaid)"]):
            ax1.text(x[i], value, f'{value}', color='black', fontsize=9, ha='center', va='bottom')

    ax2 = ax1.twinx()
    ax2.plot(x, progress["Copilot Chat H2 Incremental MAU"], marker='o', label='Copilot Chat H2 Incremental MAU', color='tab:orange', linewidth=2)
    ax2.tick_params(axis='y', labelcolor='tab:orange')

    if show_datapoint_num:
        for i, value in enumerate(progress["Copilot Chat H2 Incremental MAU"]):
            ax2.text(x[i], value, f'{value}', color='black', fontsize=9, ha='center', va='bottom')

    handles1, labels1 = ax1.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    handles = handles1 + handles2
    labels = labels1 + labels2

    ax1.legend(handles, labels, loc='upper center', bbox_to_anchor=(0.5, 1.15), ncol=2)
    plt.tight_layout()

    img_buf = io.BytesIO()
    fig.savefig(img_buf, format='png')
    img_buf.seek(0)
    plt.close(fig)

    graph = Image.open(img_buf).convert("RGBA")
    log("Account progress plot created")

    return [crop_pic(graph)]

def filtered_accounts_overview(team:pd.Series) -> tuple[pd.DataFrame, IMG]:
    """
    Generates a filtered overview of accounts for a given team member and returns the filtered DataFrame and a cropped image.
    Args:
        team (pd.Series): A pandas Series representing the team member, expected to contain an "ID" field.
    Returns:
        tuple[pd.DataFrame, IMG]: 
            - The filtered DataFrame containing only the accounts matching the filter criteria for the specified team member.
            - A cropped image (IMG) representation of the filtered DataFrame.
    Raises:
        Exception: If no filtered accounts are provided in ACCOUNTS["filter"].
    Notes:
        - The function expects global variables ACCOUNTS, COLS, DF2IMG_OPTIONS, and utility functions such as str_normalize, df2img, crop_pic, and Image to be defined elsewhere.
        - The function filters accounts based on the "TopParent" field and the provided filter list.
        - The resulting DataFrame is visualized and returned as an image after cropping.
    """

    if ACCOUNTS["filter"] is None:
        raise Exception("No filtered accounts were given")
    
    df = ACCOUNTS["data"].copy()
    df["Date"] = pd.to_datetime(df["Date"])
    
    last_date = df["Date"].max()
    id = team["ID"]

    ae_df: pd.DataFrame = df.loc[(df["AE"] == id) & (df["Date"] == last_date)]
    sum_cols = [col for col in df.columns if col not in COLS["text"] + COLS["avg"] + ["Date"]]

    ae_df.loc[:, sum_cols] = ae_df.loc[:, sum_cols].astype(int)
    COLS["sum"] = sum_cols

    filter = [str_normalize(s.lower().strip()) for s in ACCOUNTS["filter"]]
    filtered = ae_df.drop(columns=["Date", "AE"])
    filtered["TopParent"] = filtered["TopParent"].apply(
        lambda x: str_normalize(str(x).strip())
    )

    filtered = filtered[filtered["TopParent"].str.lower().str.strip().isin(filter)]
    DF2IMG_OPTIONS["df"] = filtered

    fig = df2img.plot_dataframe(**DF2IMG_OPTIONS)
    img_bytes = fig.to_image(format="png", width=1000, height=500, scale=2)
    img = Image.open(io.BytesIO(img_bytes))
    
    return filtered, crop_pic(img)

def filtered_accounts(team:pd.Series) -> list[IMG]:
    """
    Generates a filtered accounts view for a given team.
    This function processes account data for the specified team, computes summary statistics (averages, sums, and percentages)
    for relevant columns, and generates a visual representation of the results. The output is a list containing a single image
    that vertically stacks an overview picture and a cropped summary table.
    Args:
        team (pd.Series): A pandas Series representing the team for which the filtered accounts view is generated.
    Returns:
        list[IMG]: A list containing a single image object representing the stacked overview and summary table.
    """

    log("Generating filtered accounts view...", left_nl=1)

    df, overview_pic = filtered_accounts_overview(team)
    copy = df.copy()

    avg = COLS["avg"]
    sum = COLS["sum"]
    pctg = COLS["%"]

    for col in avg:
        if col in pctg:
            copy[col] = copy[col].str.rstrip("%").astype(float) / 100
            continue
        copy[col] = copy[col].astype(float) / 100

    settings = {}
    for col in avg:
        settings[col] = [
        copy[col].mean()
    ]
    for col in sum:
        settings[col] = [
        copy[col].sum()
    ]
        
    totals = pd.DataFrame(settings)
    for col in pctg:
        totals[col] = totals[col].apply(lambda x: f"{x * 100:.1f}%")

    DF2IMG_OPTIONS["df"] = totals
    fig = df2img.plot_dataframe(**DF2IMG_OPTIONS)
    img_bytes = fig.to_image(format="png", width=1000, height=500, scale=2)
    img = Image.open(io.BytesIO(img_bytes))
    
    return [stack_images_vertically(
        overview_pic, 
        crop_pic(img)
    )]
