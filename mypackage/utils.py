import time


def flatten(lst):
    return [item for sublist in lst for item in sublist]


def highlight(element):
    """Highlights (blinks) a Selenium Webdriver element."""
    driver = element._parent

    def apply_style(s):
        driver.execute_script(
            "arguments[0].setAttribute('style', arguments[1]);", element, s
        )

    original_style = element.get_attribute("style")
    apply_style("background: yellow; border: 2px solid red;")
    time.sleep(0.3)
    apply_style(original_style)
