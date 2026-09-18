import streamlit as st
import pandas as pd
import os
import random
import hashlib
import html
import re
import gspread

from datetime import datetime
from google.oauth2.service_account import Credentials


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="English-to-Telugu Translation Study",
    page_icon="🎧",
    layout="wide"
)


# ============================================================
# FILE PATHS
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_FILE = os.path.join(
    BASE_DIR,
    "25 telugu.xlsx"
)

AUDIO_DIR = os.path.join(
    BASE_DIR,
    "audio"
)


# ============================================================
# GOOGLE SHEETS HEADERS
# ============================================================

HEADERS = [
    "participant_name",
    "age_range",
    "native_language",
    "english_proficiency",
    "telugu_proficiency",
    "headphones",
    "hearing_difficulties",
    "speech_experience",
    "prosody_understanding",
    "listening_test_experience",
    "sample_id",
    "english_sentence",
    "emphasized_word",
    "audiofile",
    "prosodic_feature",
    "selected_translation",
    "selected_translation_type",
    "prosodic_rating",
    "response_time_seconds",
    "remarks",
    "last_updated"
]


# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================

def get_google_sheet():

    try:

        config = st.secrets["connections"]["gsheets"]

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]

        credentials_info = {
            "type": config["type"],
            "project_id": config["project_id"],
            "private_key_id": config["private_key_id"],
            "private_key": config["private_key"],
            "client_email": config["client_email"],
            "client_id": config["client_id"],
            "auth_uri": config["auth_uri"],
            "token_uri": config["token_uri"],
            "auth_provider_x509_cert_url":
                config["auth_provider_x509_cert_url"],
            "client_x509_cert_url":
                config["client_x509_cert_url"]
        }

        credentials = Credentials.from_service_account_info(
            credentials_info,
            scopes=scopes
        )

        client = gspread.authorize(credentials)

        spreadsheet = client.open_by_url(
            config["spreadsheet"]
        )

        worksheet_name = config.get(
            "worksheet",
            "Responses"
        )

        worksheet = spreadsheet.worksheet(
            worksheet_name
        )

        return worksheet

    except Exception as e:

        st.error(
            "Unable to connect to Google Sheets."
        )

        st.exception(e)

        return None


# ============================================================
# INITIALIZE GOOGLE SHEET
# ============================================================

def initialize_sheet():

    worksheet = get_google_sheet()

    if worksheet is None:
        return None

    try:

        existing_headers = worksheet.row_values(1)

        if not existing_headers:

            worksheet.append_row(
                HEADERS,
                value_input_option="USER_ENTERED"
            )

        else:

            if existing_headers != HEADERS:

                worksheet.update(
                    "A1:V1",
                    [HEADERS],
                    value_input_option="USER_ENTERED"
                )

        return worksheet

    except Exception as e:

        st.error(
            "Could not initialize Google Sheet."
        )

        st.exception(e)

        return worksheet


def worksheet_to_dataframe(worksheet):
    """
    Read Google Sheets using raw values instead of get_all_records().
    This avoids failures when a worksheet has duplicate/blank headers.
    """
    values = worksheet.get_all_values()

    if not values:
        return pd.DataFrame(columns=HEADERS)

    header_row = list(values[0])

    # Normalize the header row to the expected headers.
    if len(header_row) < len(HEADERS):
        header_row += [""] * (len(HEADERS) - len(header_row))
    header_row = header_row[:len(HEADERS)]

    # Use the expected headers as the dataframe schema.
    data_rows = []
    for row_values in values[1:]:
        row_values = list(row_values)
        if len(row_values) < len(HEADERS):
            row_values += [""] * (len(HEADERS) - len(row_values))
        data_rows.append(row_values[:len(HEADERS)])

    return pd.DataFrame(data_rows, columns=HEADERS)


def get_sheet_records(worksheet):
    """Return sheet rows as dictionaries without get_all_records()."""
    df = worksheet_to_dataframe(worksheet)

    if df.empty:
        return []

    return df.to_dict("records")


# ============================================================
# READ RESPONSES
# ============================================================

def read_responses():

    worksheet = initialize_sheet()

    if worksheet is None:
        return pd.DataFrame(columns=HEADERS)

    try:
        return worksheet_to_dataframe(worksheet)

    except Exception as e:

        st.error(
            "Could not read responses from Google Sheets."
        )

        st.exception(e)

        return pd.DataFrame(columns=HEADERS)


# ============================================================
# CHECK PARTICIPANT
# ============================================================

def participant_exists(participant_name):

    if not participant_name:
        return False

    df = read_responses()

    if df.empty:
        return False

    names = (
        df["participant_name"]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    return (
        participant_name.strip().lower()
        in names.values
    )


# ============================================================
# LOAD QUESTIONS FROM EXCEL
# ============================================================

@st.cache_data
def load_questions():

    if not os.path.exists(EXCEL_FILE):

        st.error(
            "Excel file not found."
        )

        st.code(EXCEL_FILE)

        st.stop()

    try:

        df = pd.read_excel(
            EXCEL_FILE
        )

    except PermissionError:

        st.error(
            """
            Permission denied while reading the Excel file.

            Please close the Excel file if it is open.
            If you are using OneDrive, make sure the file
            has finished syncing.
            """
        )

        st.code(EXCEL_FILE)

        st.info(
            "If the problem continues, move the project "
            "outside the OneDrive folder."
        )

        st.stop()

    except Exception as e:

        st.error(
            "Could not read the Excel file."
        )

        st.exception(e)

        st.stop()

    required_columns = [
        "Index",
        "Mapped English Text",
        "Telugu Text",
        "IndicTrans2 English-to-Telugu",
        "BhashaVerse English-to-Telugu",
        "google English-to-Telugu",
        "Clitic English Word",
        "Audio Path"
    ]

    missing_columns = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:

        st.error(
            "The following required columns are missing:"
        )

        st.write(
            missing_columns
        )

        st.stop()

    df = df.fillna("")

    return df


questions_df = load_questions()


# ============================================================
# GET PROSODIC FEATURE
# ============================================================

def get_prosodic_feature(row):

    possible_columns = [
        "prosodic feature",
        "prosodic_feature",
        "feature",
        "feature type",
        "feature_type",
        "prosody",
        "prosody type",
        "prosody_type"
    ]

    for col in possible_columns:

        if col in row.index:

            value = str(
                row[col]
            ).strip()

            if value:

                return value

    return ""


# ============================================================
# HIGHLIGHT EMPHASIZED WORD
# ============================================================

def highlight_emphasis(
    sentence,
    emphasized
):

    sentence = str(sentence)

    emphasized = str(
        emphasized
    ).strip()

    escaped_sentence = html.escape(
        sentence
    )

    if not emphasized:

        return escaped_sentence

    words = [
        x.strip()
        for x in emphasized
        .replace("/", ",")
        .split(",")
        if x.strip()
    ]

    words = sorted(
        words,
        key=len,
        reverse=True
    )

    for word in words:

        escaped_word = html.escape(
            word
        )

        escaped_sentence = (
            escaped_sentence.replace(
                escaped_word,
                (
                    '<span class="emphasis-word">'
                    f'{escaped_word}'
                    '</span>'
                )
            )
        )

    return escaped_sentence


# ============================================================
# STABLE RANDOMIZATION
# ============================================================

def get_randomized_options(
    participant_name,
    sample_id,
    telugu_text,
    indictrans2_translation,
    bhashaverse_translation,
    google_translation
):

    seed_string = (
        f"{participant_name.strip().lower()}_{sample_id}"
    )

    seed_hash = hashlib.sha256(
        seed_string.encode("utf-8")
    ).hexdigest()

    seed = int(seed_hash[:16], 16)

    rng = random.Random(seed)

    options = [
        {
            "text": str(telugu_text).strip(),
            "type": "Telugu Reference Translation"
        },
        {
            "text": str(indictrans2_translation).strip(),
            "type": "IndicTrans2"
        },
        {
            "text": str(bhashaverse_translation).strip(),
            "type": "BhashaVerse"
        },
        {
            "text": str(google_translation).strip(),
            "type": "Google Translation"
        },
        {
            "text": "None of these",
            "type": "none"
        }
    ]

    rng.shuffle(options)

    return options


# ============================================================
# LOAD PARTICIPANT PROGRESS
# ============================================================

def load_participant_progress(
    participant_name
):

    df = read_responses()

    if df.empty:

        return {
            "answers": {},
            "demographics": {},
            "remarks": "",
            "first_unanswered": 0
        }

    participant_rows = df[
        df["participant_name"]
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        participant_name.strip().lower()
    ]

    if participant_rows.empty:

        return {
            "answers": {},
            "demographics": {},
            "remarks": "",
            "first_unanswered": 0
        }

    answers = {}

    for _, row in participant_rows.iterrows():

        sample_id = str(
            row.get(
                "sample_id",
                ""
            )
        ).strip()

        if sample_id:

            answers[sample_id] = {

                "selected_translation":
                    row.get(
                        "selected_translation",
                        ""
                    ),

                "selected_translation_type":
                    row.get(
                        "selected_translation_type",
                        ""
                    ),

                "prosodic_rating":
                    row.get(
                        "prosodic_rating",
                        ""
                    ),

                "response_time_seconds":
                    row.get(
                        "response_time_seconds",
                        ""
                    )
            }

    first_row = participant_rows.iloc[0]

    demographics = {

        "age_range":
            first_row.get(
                "age_range",
                ""
            ),

        "native_language":
            first_row.get(
                "native_language",
                ""
            ),

        "english_proficiency":
            first_row.get(
                "english_proficiency",
                ""
            ),

        "telugu_proficiency":
            first_row.get(
                "telugu_proficiency",
                ""
            ),

        "headphones":
            first_row.get(
                "headphones",
                ""
            ),

        "hearing_difficulties":
            first_row.get(
                "hearing_difficulties",
                ""
            ),

        "speech_experience":
            first_row.get(
                "speech_experience",
                ""
            ),

        "prosody_understanding":
            first_row.get(
                "prosody_understanding",
                ""
            ),

        "listening_test_experience":
            first_row.get(
                "listening_test_experience",
                ""
            )
    }

    remarks = first_row.get(
        "remarks",
        ""
    )

    first_unanswered = len(
        questions_df
    )

    for i, row in questions_df.iterrows():

        sample_id = str(
            row["Index"]
        ).strip()

        if sample_id not in answers:

            first_unanswered = i

            break

    return {
        "answers": answers,
        "demographics": demographics,
        "remarks": remarks,
        "first_unanswered": first_unanswered
    }


# ============================================================
# CREATE RESPONSE DATA
# ============================================================

def create_response_data(
    row,
    selected_translation,
    selected_translation_type,
    prosodic_rating,
    response_time
):

    return {

        "participant_name":
            st.session_state.participant_name,

        "age_range":
            st.session_state.demographics.get(
                "age_range",
                ""
            ),

        "native_language":
            st.session_state.demographics.get(
                "native_language",
                ""
            ),

        "english_proficiency":
            st.session_state.demographics.get(
                "english_proficiency",
                ""
            ),

        "telugu_proficiency":
            st.session_state.demographics.get(
                "telugu_proficiency",
                ""
            ),

        "headphones":
            st.session_state.demographics.get(
                "headphones",
                ""
            ),

        "hearing_difficulties":
            st.session_state.demographics.get(
                "hearing_difficulties",
                ""
            ),

        "speech_experience":
            st.session_state.demographics.get(
                "speech_experience",
                ""
            ),

        "prosody_understanding":
            st.session_state.demographics.get(
                "prosody_understanding",
                ""
            ),

        "listening_test_experience":
            st.session_state.demographics.get(
                "listening_test_experience",
                ""
            ),

        "sample_id":
            str(row["Index"]),

        "english_sentence":
            str(row["Mapped English Text"]),

        "emphasized_word":
            str(row["Clitic English Word"]),

        "audiofile":
            re.split(r"[\\/]", str(row["Audio Path"]).strip())[-1],

        "prosodic_feature":
            get_prosodic_feature(row),

        "selected_translation":
            selected_translation,

        "selected_translation_type":
            selected_translation_type,

        "prosodic_rating":
            prosodic_rating,

        "response_time_seconds":
            response_time,

        "remarks":
            st.session_state.get(
                "remarks",
                ""
            ),

        "last_updated":
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
    }


# ============================================================
# SAVE RESPONSE
# ============================================================

def save_progress(
    row,
    selected_translation,
    selected_translation_type,
    prosodic_rating,
    response_time
):

    worksheet = initialize_sheet()

    if worksheet is None:
        return False

    data = create_response_data(
        row,
        selected_translation,
        selected_translation_type,
        prosodic_rating,
        response_time
    )

    try:

        records = get_sheet_records(worksheet)

        participant_name = (
            st.session_state
            .participant_name
            .strip()
            .lower()
        )

        sample_id = str(
            row["Index"]
        ).strip()

        existing_row_number = None

        for index, record in enumerate(
            records,
            start=2
        ):

            record_name = str(
                record.get(
                    "participant_name",
                    ""
                )
            ).strip().lower()

            record_sample = str(
                record.get(
                    "sample_id",
                    ""
                )
            ).strip()

            if (
                record_name == participant_name
                and record_sample == sample_id
            ):

                existing_row_number = index

                break

        row_values = [
            data.get(
                header,
                ""
            )
            for header in HEADERS
        ]

        if existing_row_number:

            worksheet.update(
                f"A{existing_row_number}:V{existing_row_number}",
                [row_values],
                value_input_option="USER_ENTERED"
            )

        else:

            worksheet.append_row(
                row_values,
                value_input_option="USER_ENTERED"
            )

        return True

    except Exception as e:

        st.error(
            "Could not save the response."
        )

        st.exception(e)

        return False


# ============================================================
# SAVE REMARKS
# ============================================================

def save_remarks():

    worksheet = initialize_sheet()

    if worksheet is None:
        return False

    try:

        records = get_sheet_records(worksheet)

        participant_name = (
            st.session_state
            .participant_name
            .strip()
            .lower()
        )

        remarks_column = (
            HEADERS.index("remarks") + 1
        )

        updated_column = (
            HEADERS.index("last_updated") + 1
        )

        for index, record in enumerate(
            records,
            start=2
        ):

            record_name = str(
                record.get(
                    "participant_name",
                    ""
                )
            ).strip().lower()

            if record_name == participant_name:

                worksheet.update_cell(
                    index,
                    remarks_column,
                    st.session_state.remarks
                )

                worksheet.update_cell(
                    index,
                    updated_column,
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )

        return True

    except Exception as e:

        st.error(
            "Could not save remarks."
        )

        st.exception(e)

        return False


# ============================================================
# SESSION STATE
# ============================================================

if "page" not in st.session_state:
    st.session_state.page = "welcome"

if "participant_name" not in st.session_state:
    st.session_state.participant_name = ""

if "current_question" not in st.session_state:
    st.session_state.current_question = 0

if "answers" not in st.session_state:
    st.session_state.answers = {}

if "randomized_options" not in st.session_state:
    st.session_state.randomized_options = {}

if "translation_selections" not in st.session_state:
    st.session_state.translation_selections = {}

if "rating_selections" not in st.session_state:
    st.session_state.rating_selections = {}

if "question_start_times" not in st.session_state:
    st.session_state.question_start_times = {}

if "remarks" not in st.session_state:
    st.session_state.remarks = ""

if "demographics" not in st.session_state:

    st.session_state.demographics = {

        "age_range": "",
        "native_language": "",
        "english_proficiency": "",
        "telugu_proficiency": "",
        "headphones": "",
        "hearing_difficulties": "",
        "speech_experience": "",
        "prosody_understanding": "",
        "listening_test_experience": ""
    }


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
<style>

.main-title {
    font-size: 42px;
    font-weight: 700;
    margin-bottom: 5px;
    color: var(--text-color);
}

.subtitle {
    font-size: 20px;
    color: var(--text-color);
    opacity: 0.65;
    margin-bottom: 35px;
}

.section-title {
    font-size: 30px;
    font-weight: 650;
    margin-top: 25px;
    margin-bottom: 15px;
    color: var(--text-color);
}

.sentence-box {
    background-color: var(--secondary-background-color);
    color: var(--text-color);
    border: 1px solid var(--text-color);
    border-radius: 12px;
    padding: 22px;
    font-size: 21px;
    line-height: 1.7;
    margin-top: 10px;
    margin-bottom: 20px;
}

.emphasis-word {
    color: #ff4b4b;
    font-weight: 800;
    text-decoration: underline;
    text-decoration-thickness: 2px;
}

.info-box {
    background-color: var(--secondary-background-color);
    color: var(--text-color);
    border: 1px solid var(--text-color);
    border-radius: 12px;
    padding: 20px;
    margin-top: 15px;
    margin-bottom: 20px;
}

.info-box p,
.info-box strong,
.info-box b {
    color: var(--text-color) !important;
}

div[class*="st-key-translation_choice_"] {
    width: 100% !important;
}

div[class*="st-key-translation_choice_"] div[role="radiogroup"] {
    display: flex !important;
    flex-direction: column !important;
    width: 100% !important;
    gap: 30px !important;
    margin-top: 15px !important;
    margin-bottom: 20px !important;
}

div[class*="st-key-translation_choice_"] div[role="radiogroup"] label {
    display: flex !important;
    flex-direction: row !important;
    align-items: flex-start !important;
    width: 100% !important;
    min-height: 115px !important;
    box-sizing: border-box !important;
    padding: 26px 30px !important;
    margin: 0 !important;
    border: 2px solid var(--text-color) !important;
    border-radius: 15px !important;
    background-color: var(--secondary-background-color) !important;
    color: var(--text-color) !important;
    cursor: pointer !important;
    white-space: normal !important;
    overflow: visible !important;
    transition: background-color 0.2s ease, border-color 0.2s ease, box-shadow 0.2s ease, transform 0.2s ease !important;
}

div[class*="st-key-translation_choice_"] div[role="radiogroup"] label p {
    display: block !important;
    width: 100% !important;
    max-width: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
    font-size: 21px !important;
    line-height: 1.8 !important;
    color: var(--text-color) !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
    word-break: normal !important;
    overflow-wrap: anywhere !important;
}

div[class*="st-key-translation_choice_"] div[role="radiogroup"] label > div {
    width: 100% !important;
    max-width: 100% !important;
    white-space: normal !important;
    overflow: visible !important;
    color: var(--text-color) !important;
}

div[class*="st-key-translation_choice_"] div[role="radiogroup"] label input {
    flex: 0 0 auto !important;
    width: 20px !important;
    height: 20px !important;
    margin-top: 5px !important;
    margin-right: 18px !important;
}

div[class*="st-key-translation_choice_"] div[role="radiogroup"] label:hover {
    background-color: var(--secondary-background-color) !important;
    border-color: #4da3ff !important;
    box-shadow: 0 5px 18px rgba(77, 163, 255, 0.20) !important;
    transform: translateY(-2px) !important;
}

div[class*="st-key-translation_choice_"] div[role="radiogroup"] label:has(input:checked) {
    background-color: var(--secondary-background-color) !important;
    border: 2px solid #4da3ff !important;
    box-shadow: 0 0 0 1px #4da3ff, 0 6px 20px rgba(77, 163, 255, 0.20) !important;
}

audio {
    width: 100% !important;
    margin-top: 10px;
    margin-bottom: 20px;
}

@media (max-width: 768px) {
    div[class*="st-key-translation_choice_"] div[role="radiogroup"] {
        gap: 20px !important;
    }
    div[class*="st-key-translation_choice_"] div[role="radiogroup"] label {
        min-height: 100px !important;
        padding: 20px 22px !important;
    }
    div[class*="st-key-translation_choice_"] div[role="radiogroup"] label p {
        font-size: 18px !important;
        line-height: 1.7 !important;
    }
}

</style>
""",
    unsafe_allow_html=True
)


# ============================================================
# WELCOME PAGE
# ============================================================

if st.session_state.page == "welcome":

    st.markdown(
        '<div class="main-title">'
        'English-to-Telugu Translation Study'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        "## About the Study"
    )

    st.write(
        """
        This study is designed to examine how prosodic features
        in English speech are reflected in English-to-Telugu
        translation.

        In particular, the study focuses on two aspects of
        prosody: **emphasis** and **rising contour**.

        You will listen to short English speech recordings,
        rate how strongly you perceive the relevant prosodic
        feature, and then choose the Telugu translation that
        best matches the intended meaning and prosodic
        interpretation.
        
        **1. Listen carefully to the English audio.**

        You may replay the recording as many times as necessary.

        **2. Pay attention to the speaker's prosody**, especially:

        - **Emphasis:** a word or phrase may sound more prominent
          than the surrounding words.
        - **Rising contour:** the pitch may rise toward the end
          of a word, phrase, or sentence.

        **3. Pay attention to the indicated word or phrase**
        where applicable.

        **4. Rate how strongly you perceive the relevant
        prosodic feature in the audio.**

        **5. Choose the Telugu translation** that best matches
        the intended meaning and the prosodic interpretation
        you perceived.
        """
    )


    st.markdown(
        """
        <div class="info-box">

        <b>Important:</b>

        Your responses are based on your perception of the
        audio. There are no right or wrong answers from the
        participant's perspective.

        Note: Only focus on the given emphasized word.
        There might be other emphasized words in the audio , just ignore those.

        </div>
        """,
        unsafe_allow_html=True
    )

    if st.button(
        "Begin Study →",
        type="primary",
        use_container_width=True
    ):

        st.session_state.page = "participant_info"

        st.rerun()


# ============================================================
# PARTICIPANT INFORMATION
# ============================================================

elif st.session_state.page == "participant_info":

    st.markdown(
        '<div class="section-title">'
        'Participant Information'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "Please provide the following information before beginning the study."
    )

    participant_name = st.text_input(
        "Participant name",
        value=st.session_state.participant_name,
        placeholder="Enter your name"
    )

    age_range = st.radio(
        "Age range",
        [
            "Below 18",
            "18–24",
            "25–34",
            "35–44",
            "45–54",
            "55 or above"
        ],
        index=None
    )

    native_language = st.text_input(
        "Native language(s)"
    )

    english_proficiency = st.radio(
        "English proficiency",
        [
            "Beginner",
            "Intermediate",
            "Advanced",
            "Native / Near-native"
        ],
        index=None
    )

    telugu_proficiency = st.radio(
        "Telugu proficiency",
        [
            "None",
            "Beginner",
            "Intermediate",
            "Advanced",
            "Native / Near-native"
        ],
        index=None
    )

    headphones = st.radio(
        "Are you using headphones or earphones?",
        [
            "Yes",
            "No"
        ],
        index=None
    )

    hearing_difficulties = st.radio(
        "Do you have any difficulty hearing speech?",
        [
            "Yes",
            "No",
            "Prefer not to say"
        ],
        index=None
    )

    speech_experience = st.radio(
        "Do you have previous experience with speech, "
        "linguistics, audio, or related research?",
        [
            "Yes",
            "No"
        ],
        index=None
    )

    prosody_understanding = st.radio(
        "How familiar are you with the concept of prosody?",
        [
            "Not familiar",
            "Slightly familiar",
            "Moderately familiar",
            "Very familiar"
        ],
        index=None
    )

    listening_test_experience = st.radio(
        "Have you participated in a listening test before?",
        [
            "Yes",
            "No"
        ],
        index=None
    )

    if st.button(
        "Continue →",
        type="primary",
        use_container_width=True
    ):

        if not participant_name.strip():

            st.warning(
                "Please enter your name."
            )

        elif not age_range:

            st.warning(
                "Please select your age range."
            )

        elif not native_language.strip():

            st.warning(
                "Please enter your native language."
            )

        elif not english_proficiency:

            st.warning(
                "Please select your English proficiency."
            )

        elif not telugu_proficiency:

            st.warning(
                "Please select your Telugu proficiency."
            )

        elif not headphones:

            st.warning(
                "Please indicate whether you are using headphones or earphones."
            )

        elif not hearing_difficulties:

            st.warning(
                "Please answer the hearing-difficulty question."
            )

        elif not speech_experience:

            st.warning(
                "Please answer the speech/audio experience question."
            )

        elif not prosody_understanding:

            st.warning(
                "Please select your familiarity with prosody."
            )

        elif not listening_test_experience:

            st.warning(
                "Please answer the listening-test experience question."
            )

        else:

            st.session_state.participant_name = (
                participant_name.strip()
            )

            st.session_state.demographics = {

                "age_range":
                    age_range,

                "native_language":
                    native_language.strip(),

                "english_proficiency":
                    english_proficiency,

                "telugu_proficiency":
                    telugu_proficiency,

                "headphones":
                    headphones,

                "hearing_difficulties":
                    hearing_difficulties,

                "speech_experience":
                    speech_experience,

                "prosody_understanding":
                    prosody_understanding,

                "listening_test_experience":
                    listening_test_experience
            }

            progress = load_participant_progress(
                st.session_state.participant_name
            )

            if participant_exists(
                st.session_state.participant_name
            ):

                st.session_state.answers = (
                    progress["answers"]
                )

                st.session_state.remarks = (
                    progress["remarks"]
                )

                st.session_state.current_question = (
                    progress["first_unanswered"]
                )

                if (
                    st.session_state.current_question
                    >= len(questions_df)
                ):

                    st.session_state.page = "completed"

                else:

                    st.session_state.page = "instructions"

            else:

                st.session_state.current_question = 0

                st.session_state.page = "instructions"

            st.rerun()


# ============================================================
# INSTRUCTIONS
# ============================================================

elif st.session_state.page == "instructions":

    st.markdown(
        """
        ### Prosodic perception rating

        | Rating | Meaning |
        |---|---|
        | **1** | Not perceived at all |
        | **2** | Slightly perceived |
        | **3** | Moderately perceived |
        | **4** | Strongly perceived |
        | **5** | Very strongly perceived |
        """
    )

    if st.button(
        "Start Experiment →",
        type="primary",
        use_container_width=True
    ):

        st.session_state.page = "experiment"

        st.rerun()


# ============================================================
# EXPERIMENT
# ============================================================

elif st.session_state.page == "experiment":

    current_index = (
        st.session_state.current_question
    )

    total_questions = len(
        questions_df
    )

    if current_index >= total_questions:

        st.session_state.page = "remarks"

        st.rerun()

    row = questions_df.iloc[
        current_index
    ]

    sample_id = str(
        row["Index"]
    ).strip()

    english_sentence = str(
        row["Mapped English Text"]
    ).strip()

    telugu_text = str(
        row["Telugu Text"]
    ).strip()

    indictrans2_translation = str(
        row["IndicTrans2 English-to-Telugu"]
    ).strip()

    bhashaverse_translation = str(
        row["BhashaVerse English-to-Telugu"]
    ).strip()

    google_translation = str(
        row["google English-to-Telugu"]
    ).strip()

    emphasized_word = str(
        row["Clitic English Word"]
    ).strip()

    audio_source = str(
    row["Audio Path"]
    ).strip()

    audiofile = re.split(
    r"[\\/]",
    audio_source
    )[-1]

    prosodic_feature = get_prosodic_feature(
        row
    )


    # ========================================================
    # PROGRESS
    # ========================================================

    st.progress(
        current_index / total_questions
    )

    st.write(
        f"Question {current_index + 1} of {total_questions}"
    )

    st.markdown("---")


    # ========================================================
    # ENGLISH SENTENCE
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        'English sentence'
        '</div>',
        unsafe_allow_html=True
    )

    highlighted_sentence = highlight_emphasis(
        english_sentence,
        emphasized_word
    )

    st.markdown(
        f"""
        <div class="sentence-box">
            {highlighted_sentence}
        </div>
        """,
        unsafe_allow_html=True
    )

    if prosodic_feature:

        st.info(
            f"Prosodic feature: **{prosodic_feature}**"
        )

    if emphasized_word:

        st.write(
            f"**Indicated word(s):** {emphasized_word}"
        )


    # ========================================================
    # AUDIO
    # ========================================================

    st.markdown(
        '<div class="section-title">'
        'Listen to the audio'
        '</div>',
        unsafe_allow_html=True
    )

    audio_path = os.path.join(
        AUDIO_DIR,
        audiofile
    )

    if os.path.exists(audio_path):

        st.audio(
            audio_path
        )

    else:

        st.error(
            f"Audio file not found: {audiofile}"
        )


    # ========================================================
    # TIMER
    # ========================================================

    if sample_id not in st.session_state.question_start_times:

        st.session_state.question_start_times[
            sample_id
        ] = datetime.now()


    # ========================================================
    # PROSODIC RATING
    # ========================================================

    st.markdown("---")

    st.markdown(
        '<div class="section-title">'
        'Rate your perception'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "How strongly did you perceive the relevant "
        "prosodic feature in the audio?"
    )

    st.caption(
        "Please base your rating on what you hear "
        "in the audio recording."
    )

    rating_options = [

        "1 — Not perceived at all",

        "2 — Slightly perceived",

        "3 — Moderately perceived",

        "4 — Strongly perceived",

        "5 — Very strongly perceived"
    ]

    previous_rating = (
        st.session_state
        .rating_selections
        .get(
            sample_id,
            None
        )
    )

    rating_index = None

    if previous_rating in rating_options:

        rating_index = (
            rating_options.index(
                previous_rating
            )
        )

    selected_rating = st.radio(

        "Prosodic feature rating",

        rating_options,

        index=rating_index,

        key=f"rating_{sample_id}"
    )


    # ========================================================
    # TRANSLATION
    # ========================================================

    st.markdown("---")

    st.markdown(
        '<div class="section-title">'
        'Choose the Telugu translation'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        "Four Telugu versions are shown below. Select the "
        "version that best matches the intended meaning and "
        "prosodic interpretation."
    )


    # ========================================================
    # ENSURE FOUR TRANSLATION OPTIONS
    # ========================================================

    if not telugu_text:
        st.error(
            "Telugu reference translation is missing for this question."
        )
        st.stop()

    if not indictrans2_translation:
        st.error(
            "IndicTrans2 translation is missing for this question."
        )
        st.stop()

    if not bhashaverse_translation:
        st.error(
            "BhashaVerse translation is missing for this question."
        )
        st.stop()

    if not google_translation:
        st.error(
            "Google translation is missing for this question."
        )
        st.stop()


    # ========================================================
    # RANDOMIZE FIVE OPTIONS INCLUDING NONE OF THESE
    # ========================================================

    options = get_randomized_options(

        st.session_state.participant_name,

        sample_id,

        telugu_text,

        indictrans2_translation,

        bhashaverse_translation,

        google_translation
    )

    st.session_state.randomized_options[
        sample_id
    ] = options

    display_texts = [

        option["text"]

        for option in options
    ]


    # ========================================================
    # PREVIOUS SELECTION
    # ========================================================

    previous_translation = (

        st.session_state
        .translation_selections
        .get(
            sample_id,
            None
        )
    )

    translation_index = None

    if previous_translation in display_texts:

        translation_index = (
            display_texts.index(
                previous_translation
            )
        )


    # ========================================================
    # TRANSLATION OPTION BOXES
    # ========================================================

    selected_translation = st.radio(

        "Translation options",

        display_texts,

        index=translation_index,

        key=f"translation_choice_{sample_id}",

        label_visibility="collapsed"
    )


    # ========================================================
    # NAVIGATION
    # ========================================================

    st.markdown(
        "<br>",
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(
        [1, 1]
    )

    with col1:

        if current_index > 0:

            if st.button(
                "← Previous",
                use_container_width=True
            ):

                st.session_state.current_question -= 1

                st.rerun()

    with col2:

        button_text = (

            "Submit Study"

            if current_index == total_questions - 1

            else "Next →"
        )

        if st.button(

            button_text,

            type="primary",

            use_container_width=True
        ):

            # ------------------------------------------------
            # VALIDATE RATING
            # ------------------------------------------------

            if not selected_rating:

                st.warning(
                    "Please provide a prosodic feature rating."
                )

                st.stop()


            # ------------------------------------------------
            # VALIDATE TRANSLATION
            # ------------------------------------------------

            if not selected_translation:

                st.warning(
                    "Please select a Telugu translation."
                )

                st.stop()


            # ------------------------------------------------
            # FIND TRANSLATION TYPE
            # ------------------------------------------------

            selected_type = ""

            for option in options:

                if (
                    option["text"]
                    == selected_translation
                ):

                    selected_type = (
                        option["type"]
                    )

                    break


            # ------------------------------------------------
            # RESPONSE TIME
            # ------------------------------------------------

            start_time = (

                st.session_state
                .question_start_times
                .get(
                    sample_id,
                    datetime.now()
                )
            )

            response_time = (

                datetime.now()
                - start_time
            ).total_seconds()


            # ------------------------------------------------
            # SAVE SESSION DATA
            # ------------------------------------------------

            st.session_state.rating_selections[
                sample_id
            ] = selected_rating

            st.session_state.translation_selections[
                sample_id
            ] = selected_translation

            st.session_state.answers[
                sample_id
            ] = {

                "selected_translation":
                    selected_translation,

                "selected_translation_type":
                    selected_type,

                "prosodic_rating":
                    selected_rating,

                "response_time_seconds":
                    response_time
            }


            # ------------------------------------------------
            # SAVE GOOGLE SHEETS
            # ------------------------------------------------

            success = save_progress(

                row,

                selected_translation,

                selected_type,

                selected_rating,

                response_time
            )

            if success:

                if (
                    current_index
                    ==
                    total_questions - 1
                ):

                    st.session_state.current_question += 1

                    st.session_state.page = "remarks"

                else:

                    st.session_state.current_question += 1

                st.rerun()


# ============================================================
# REMARKS
# ============================================================

elif st.session_state.page == "remarks":

    st.markdown(
        '<div class="section-title">'
        'Thank You for Completing the Study'
        '</div>',
        unsafe_allow_html=True
    )

    st.write(
        """
        You have completed all the listening and translation
        questions.

        Before submitting, please provide any comments or
        remarks you may have about the study.
        """
    )

    st.write(
        """
        You may comment on:

        - Clarity of the instructions
        - Understanding the question
        - Understanding translations
        - Emphasis understanding
        - Rising contour understanding
        - Difficulty of the task
        - Anything else you would like to mention
        """
    )

    remarks = st.text_area(

        "Please enter your comments or remarks",

        value=st.session_state.remarks,

        height=180,

        placeholder="Enter your remarks here..."
    )

    st.session_state.remarks = remarks

    if st.button(

        "Submit Remarks",

        type="primary",

        use_container_width=True
    ):

        if save_remarks():

            st.session_state.page = "completed"

            st.rerun()


# ============================================================
# COMPLETED
# ============================================================

elif st.session_state.page == "completed":

    st.markdown(
        '<div class="main-title">'
        'Study Completed'
        '</div>',
        unsafe_allow_html=True
    )

    st.success(
        "Thank you for participating in the study!"
    )

    st.write(
        """
        Your responses and remarks have been recorded
        successfully.

        You may now close this page.
        """
    )