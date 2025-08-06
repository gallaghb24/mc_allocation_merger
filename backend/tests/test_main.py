"""Basic tests for the Streamlit app module."""


def test_import_streamlit_app() -> None:
    """Ensure the Streamlit app can be imported."""
    import frontend.streamlit_app  # noqa: F401
