class MyLocators:
    # declare some elements
    url_input = "[name='url']"
    analyze_btn = "//div[@class='analyze-cell']"
    mobile_result = "//div[@class='tab-title tab-mobile']"
    desktop_result = "//div[@class='tab-title tab-desktop']"
    scores_performance = "//div[@class='lh-gauge__percentage']"
    my_header = {"user-agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                 'Chrome/93.0.4577.82 Safari/537.36'}
    pagespeed_url = r'https://developers.google.com/speed/pagespeed/insights/'
