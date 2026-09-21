# Run: python test_classify.py  — pins the order-sensitive TYPE_RULES cases.
from sharesansar_events_and_mail import classify_event

CASES = {
    "Book Closure for 17th AGM and Cash Dividend [XYZ]": "Book Closure – AGM",
    "4th Annual General Meeting of RBCL [RBCL]": "AGM – Meeting Day",
    "NIFRA_AGM [NIFRA]": "AGM",
    "नवौं बार्षिक साधारण सभा सम्बन्धी सूचना [PURE]": "AGM – Notice",
    "BPCL Calls EGM [BPCL]": "SGM – Notice",
    "Opening Day of auction of X Life Insurance for 31,588 units promoter shares.": "Auction",
    "Opening Day of ratio 1:1 Right Share of X Hydropower Limited.": "Right Share",
    'Opening Day of (5,00,000 units @ Rs.1000 per unit) "8% X Debenture 2089"': "Debenture/Bond",
    "Closing Day of 1,00,000 units of X Hydropower Limited to the Foreign Nepalese Immigrants.": "IPO",
    "Dividend Declaration of Preference Share [MBLPNP]": "Dividend – Declaration",
    "Information of termination of CEO of Himalayan Investment Banker Ltd. [HLICF]": "Board/Management",
    "Information Regarding Investment [NRN]": "Investment/Project",
    "मर्जर प्रयोजनका लागि धितोपत्रको कारोबार रोक्का राखिएको बारे": "Merger/Acquisition",
    "आधारभुत शेयरधनीको सेयर रोक्का सम्बन्धमा [WNLB]": "Share Freeze/Pledge",
    "Auditor committee Reform/Company Secretary [SIPD]": "Board/Management",
    "बाह्य लेखापरीक्षक नियुक्तिको जानकारी सम्बन्धमा । [HLBSL]": "Auditor",
    "MDB [MDB]": "General Notice",
    # Subtypes
    "12TH AGM EXTENTION LETTER [MKJC]": "AGM – Postponed/Changed",
    "13th Agm Minute [MERO]": "AGM – Minutes/Decisions",
    "Successful Completion of 13th, 14th and 15th Annual General Meeting of Greenlife Hydropower Ltd. [GLH]": "AGM – Completed",
    "AGM [BJHL]": "AGM",
    "Samudayik Laghubitta decleared no dividend in fy 2081-82 [SLBSL]": "Dividend – No Dividend",
    "Dividend Decleration [SANIMA]": "Dividend – Declaration",
    "SMHL Bonus Share declaration [SMHL]": "Dividend – Bonus Share",
    "Book Closure Date for SGM of Union Hydropower Limited.": "Book Closure – SGM",
    "Book Close notice of RBB Mutual Fund 1 (RMF1)": "Book Closure – Mutual Fund",
    "Book Close [DOLTI]": "Book Closure",
    "Something entirely new [ABC]": "Other",
}

if __name__ == "__main__":
    bad = [(t, want, classify_event(t)) for t, want in CASES.items() if classify_event(t) != want]
    for t, want, got in bad:
        print(f"FAIL {t!r}: want {want}, got {got}")
    assert not bad
    print(f"ok ({len(CASES)} cases)")
