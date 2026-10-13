export const SELECTORS = {
  LOGIN: {
    USERNAME_INPUT: '#id_username, input[name="username"], #username',
    PASSWORD_INPUT: '#id_password, input[name="password"], #password',
    SUBMIT_BUTTON: 'input[type="submit"], button[type="submit"], #login-btn, input[value="Log in"]',
  },
  DASHBOARD: {
    SUBJECTS_MENU: 'text=Subjects',
    SUBJECTS_LINK: 'a[href*="/school_module/subject/"]',
  },
  SUBJECT: {
    ADD_NEW_BUTTON: 'a:has-text("Add subject"), button:has-text("Add New Subject"), .addlink',
    BOARD_DROPDOWN: '#id_education_board, select[name="education_board"], select[name*="board"], #board-select',
    GRADE_DROPDOWN: '#id_grade, select[name="grade"], #grade-select',
    NAME_INPUT: '#id_name, input[name="name"], #subject-name',
    SHORT_NAME_INPUT: '#id_short_name, input[name="short_name"], #short-name',
    CODE_INPUT: '#id_code, input[name="code"], #subject-code',
    DESCRIPTION_INPUT: '#id_description, textarea[name="description"], #subject-description',
    SAVE_BUTTON: 'input[name="_save"], button[name="_save"], button:has-text("SAVE")',
    SAVE_AND_ADD_ANOTHER_BUTTON: 'input[name="_addanother"], button[name="_addanother"], input[value="Save and add another"], button:has-text("Save and add another")',
    SAVE_AND_CONTINUE_EDITING_BUTTON: 'input[name="_continue"], button[name="_continue"], button:has-text("Save and continue editing")',
    SUCCESS_MESSAGE: '.success, .messagelist .success, .toast-success',
    ERROR_MESSAGE: '.errorlist, .errors, .toast-error',
    DUPLICATE_MESSAGE: '.errorlist, .errors, text="already exists"',
    POPUP_DIALOG: '.modal-dialog',
    CLOSE_POPUP: 'button.close-modal',
  }
};
