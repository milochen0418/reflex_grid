"""UserForms on screen: a local React component (``assets/grid/userform.jsx``), a plain asset like
the grid (see :mod:`.sheet_view`), drawing the forms a person's macros show from the engine's
snapshots and sending back what the person does with them (``macro.formEvent``)."""

import reflex as rx

from ..state import GridState

_JS = rx.asset("grid/userform.jsx")


class UserForms(rx.Component):
    library = _JS.importable_path
    tag = "UserForms"

    forms: rx.Var[list]
    # Pictures (data URLs) by the key the forms' snapshots carry instead.
    pictures: rx.Var[dict]
    # The form a macro waits on (a modal one): on top, the others out of reach.
    waiting: rx.Var[int]
    # A MsgBox, an InputBox or a busy macro is in front: nothing on the forms can be used.
    blocked: rx.Var[bool]
    # The VBA editor's picture of a form's design.
    design: rx.Var[bool]
    on_event: rx.EventHandler[lambda e: [e]]


user_forms = UserForms.create


def forms_layer() -> rx.Component:
    """The person's forms, over the workbook."""
    return user_forms(
        forms=GridState.vba_forms,
        pictures=GridState.vba_form_pictures,
        waiting=GridState.vba_form_waiting,
        blocked=GridState.vba_forms_blocked,
        design=False,
        on_event=GridState.form_event,
        id="userforms",
    )


# The designer imports the form drawing from userform.jsx, served beside it.
_DESIGNER_JS = rx.asset("grid/form_designer.jsx")


class FormDesigner(rx.Component):
    """The VBA editor's form designer (``assets/grid/form_designer.jsx``)."""

    library = _DESIGNER_JS.importable_path
    tag = "FormDesigner"

    form: rx.Var[dict]
    selected: rx.Var[str]
    tool: rx.Var[str]
    pictures: rx.Var[dict]
    on_event: rx.EventHandler[lambda e: [e]]


form_designer = FormDesigner.create
