package com.tranquyet.util.screen;

/**
 * Screen: common information which is used to show for client, such as: options and error messages
 */
public class Screen {
    /**
     * show main screen of application
     */
    public static void showScreen() {
        System.out.println("");
        System.out.println("-----------Manage Student------------");
        System.out.println("Choose the option:");
        System.out.println("1. Add Student");
        System.out.println("2. Show Student");
        System.out.println("3. Create Excel file");
        System.out.println("4. Create demo Excel file: ");
        System.out.println("5. Exit;");
        System.out.println("Enter the option: ");
    }

    /**
     * show message to client
     *
     * @param message contain the information of error actions
     */
    public static void showError(String message) {
        System.err.println(message);
    }

    /**
     * show successful message when create Excel file to client
     *
     * @param message contain information of successful actions
     */
    public static void showSuccess(String message) {
        System.out.println(message);
    }
}
