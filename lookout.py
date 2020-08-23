RED = "\033[;31m"
BLUE = "\033[34m"
GREEN = "\033[32m"
MAGENTA = "\033[35m"
RESET = "\033[;0m"

try:
    import requests
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    import smtplib
    from email.mime.text import MIMEText
    import time
    import msvcrt
except Exception as error:
    print(MAGENTA + "LOOKOUT >" + RED + " FATAL: " + error)
    exit

class Lookout():
    """A class for scraping websites and sending updates via an SMTP server."""

    def __init__(self):
        self.to = None
        self.fromm = None
        self.email = None
        self.emailPassword = None
        self._smtpServer = "smtp.gmail.com"
        self._smtpServerPort = 587
        self._server = None
        self._driverOptions = None
        self._driver = None
        self.driverAddress = None
        self.exitCondition = None
        self.exitConditionSearchType = None
        self.exitConditionSearchElement = None
        self.sendMessage = None
        self.conds = []
        self.stepArgs = []

    def Init(self):
        """Initializes the SMTP server and selenium driver.

        This is not included in __init__ becuase the result must be returned.
        The SMTP server is closed after opening to make sure it can be opended,
        but is closed so that it is not open until needed.
        
        Returns:
            An error if an error has occurred, otherwise 0.
        """

        print(MAGENTA + "LOOKOUT >" + RESET + " Initializing...")
        error = self._InitSMTPServer()
        if error != 0: return error
        self.QuitSMTPServer()
        error = self._InitDriver()
        if error != 0: return error
        else:
            print(MAGENTA + "LOOKOUT >" + RESET + " Initialization successfull")
            return 0

    #=============================SMTP SERVER==================================

    def _InitSMTPServer(self):
        """Inits the SMTP server.

        Initializes the SMTP server with the specified class members
        self.smtpServer, self.smtpServerPort, self.email, and self.emailPassword.

        Returns:
            An error if an error has occurred, otherwise 0.
        """

        print(BLUE + "SMTP SERVER >" + RESET + " Initializing...")
        try:
            self._server = smtplib.SMTP(self._smtpServer, self._smtpServerPort)
            self._server.starttls()
            self._server.login(self.email, self.emailPassword)
        except Exception as error: return error
        else:
            print(BLUE + "SMTP SERVER >" + RESET + " Initialization successfull")
            return 0

    def QuitSMTPServer(self):
        """Quits the SMTP server"""

        print(BLUE + "SMTP SERVER >" + RESET + " Quitting")
        self._server.quit()

    def SendUpdate(self):
        """Sends the defined message
        
        Sends self.sendMessage via the SMTP server to self.to from self.fromm.
        """

        try:
            self._InitSMTPServer()
            message = MIMEText(self.sendMessage)
            message["From"] = self.fromm
            message["To"] = self.to
            self._server.sendmail(self.email, self.to, message.as_string())
            print(BLUE + "SMTP SERVER >" + RESET + " Sent message \"" + self.sendMessage + "\" from " + self.fromm + " to " + self.to)
        except Exception as error: print("SMTP SERVER >" + RESET + " " + error)

    def SetSMTPDetails(self, to_, from_, email_, password_):
        """Sets the details for the SMTP server.

        Args:
            to_:
                The address to send the message to.
            from_:
                The identifier of the sender.
            email_:
                The email used to login to the SMTP server.
            password_:
                The email password used to login to the SMTP server.
        """

        self.to = to_
        self.fromm = from_
        self.email = email_
        self.emailPassword = password_

    def _SetSMTPServerDetails(self, address, port):
        """Sets the SMTP server detials.

        This should be changed only if the server is not Gmail.

        Args:
            addresss:
                The SMTP server addresss to set.
            port:
                The SMTP server port to set.
        """

        self._smtpServer = address
        self._smtpServerPort = port

    def SetSMTPSendMessage(self, message):
        """Sets the driver send message."""

        self.sendMessage = message

    #===========================END SMTP SERVER================================

    #=============================CHROMEDRIVER=================================

    def _InitDriver(self):
        """Initis the selenium web driver.

        Initializes the web driver, which is chromium. It creates an Options()
        object for passing parameters to chromium when the process is spawned.
        The parameters for chromium are headless and log-level=3.  log-level=3
        makes chromium not output info to stdout.

        Returns:
            An error if an error has occurred, otherwise 0.
        """

        print(GREEN + "DRIVER >" + RESET + " Initializing...")
        try:
            self._driverOptions = Options()
            self._driverOptions.add_argument("--headless")
            self._driverOptions.add_argument("--log-level=3")
            self._driver = webdriver.Chrome(options=self._driverOptions)
        except Exception as error: return error
        else:
            print(GREEN + "DRIVER >" + RESET + " Initialization successfull")
            return 0

    def QuitDriver(self):
        """Quits the driver"""

        print(GREEN + "DRIVER >" + RESET + " Quitting")
        self._driver.quit()

    def DriverLoad(self, address):
        """The driver loads the webpage.

        Args:
            address:
                The address to load.
        """
        print(GREEN + "DRIVER > " + RESET + "Attempting to load address " + GREEN + address)
        self._driver.get(address)

    def FindFirstElement(self, searchMethod, element):
        """Finds elements with the specified search terms.

        Creates a list of all elements found using the search terms; first
        searchMethod determines how the driver will find the elements. Options
        are xpath and id. element is then passed as a parameter to either
        find_elements_by_xpath or find_elements_by_id respectively, which
        returns a list of the elments found.

        Args:
            searchMethod: A string determining which function will be used to
                return a list of elements. Must be either "xpath" or "id".
            element: A string passed with the function used to find the
                elements.

        Returns:
            The first element in the list of elements found, or None if the
            returned list is empty.
        """

        print(GREEN + "DRIVER >" + RESET + " Searching for first element by " + searchMethod + GREEN + " " + element)
        elements = None
        if searchMethod == "xpath":
            elements = self._driver.find_elements_by_xpath(element)
        else:
            elements = self._driver.find_elements_by_id(element)
        if len(elements) > 0:
            print(GREEN + "DRIVER >" + RESET + " Found element")
            return elements[0]
        else:
            print(GREEN + "DRIVER >" + RESET + " Element not found")
            return None

    def SetDriverAddress(self, address):
        """Sets the address the driver connects to during Execute."""
        
        self.driverAddress = address

    def _IsInvalidSearchMethod(self, method):
        """Determines whether method is a valid search method or not.

        Valid search methods are "xpath" and "id".

        Returns:
            A boolean: whether the method is valid.
        """

        return method != "xpath" and method != "id"

    def SetDriverSendCondition(self, cond, searchType, element):
        """Sets the condition in which the driver will send the update.

        If the current step is "check send" in Execute(), the condition
        specified via SetDriverSendCondition is evaluated. If it is met,
        then the message is sent.

        Raises:
            Exception: Parameters are invalid.
        """

        if cond != "found" and cond != "not found" and cond != "changed":
            raise Exception(GREEN + "DRIVER >" + RED + " Invalid condition given to SetDriverSendCondtion.  Valid conditions are: \"found\", \"not found\", or \"changed\". Given: " + ("\"" + cond + "\"") if cond != None else "\"NoneType\"")
        if self._IsInvalidSearchMethod(searchType):
            raise Exception(GREEN + "DRIVER >" + RED + " Invalid search type given to SetDriverSendCondition")
        self.exitCondition = cond
        self.exitConditionSearchType = searchType
        self.exitConditionSearchElement = element

    #===========================END CHROMEDRIVER===============================

    def Cleanup(self):
        """Cleans up Lookout.
        
        Quits chromedriver and the SMTP server.
        """

        self.QuitDriver()
        self.QuitSMTPServer()

    def AddStep(self, cond, *kwargs):
        """Adds a step to self.stepArgs and self.conds for use in Execute().

        Attempts to add a step to self.conds and self.stepArgs, provided the
        parameters are valid.

        Args:
            cond:
                A condition or action that is taken during Execute(). Valid
                options are "login", "wait for element appear",
                "wait for seconds", "exit", and "check send".  Each cond has
                its own condition which must be met for the step to be added.
            *kwargs:
                A list containing an arbitrary number of argumets.  If cond is
                "login", this parameter must be
                [{username element search method},
                {username element search path}, {username},
                {password element search method},
                {password element search path}, {password},
                {submit buttom element search method},
                {submit button seach path}]. If cond is "wait for element to
                appear" or "wait for element to dissapear", this parameter
                must be [{element seach method}, {element seach path}]. If cond
                is "wait for seconds", this parameter must be
                [{time to sleep in seconds}].
            
            Returns:
                An error if an error has occurred, otherwise 0.
        """

        try:
            if cond == "login":
                if len(kwargs) != 8:
                    raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Invalid number of args given to AddStep with specified cond as \"login\". Expected 9, given: " + str(len(kwargs)))
                if self._IsInvalidSearchMethod(kwargs[0]):
                    raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Invalid second argument given to AddStep. Second argument must be: \"xpath\" or \"id\". Given: \"" + kwargs[0] + "\"")
                if self._IsInvalidSearchMethod(kwargs[3]):
                    raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Invalid fifth argument given to AddStep. Second argument must be: \"xpath\" or \"id\". Given: \"" + kwargs[0] + "\"")
                if self._IsInvalidSearchMethod(kwargs[6]):
                    raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Invalid eighth argument given to AddStep. Seventh argument must be: \"xpath\" or \"id\". Given: \"" + kwargs[0] + "\"")
                self.conds.append(cond)
                self.stepArgs.append(kwargs)
            elif cond == "wait for element appear" or cond == "wait for element dissapear":
                if len(kwargs) != 2:
                    raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Invalid number of aruments given to AddStep. Specified cond requires there to be two more arguments given: a search method and a search string")
                if kwargs[0] == "xpath" or kwargs[0] == "id":
                    raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Invalid second argument given to AddStep. Second argument must be: \"xpath\" or \"id\". Given: \"" + kwargs[0] + "\"")
                self.conds.append(cond)
                self.stepArgs.append(kwargs)
            elif cond == "wait for seconds":
                if len(kwargs) != 1:
                    raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Invalid number of args given to AddStep with specified cond as \"wait for seconds\". Expected 1, given: " + str(len(kwargs)))
                self.conds.append(cond)
                self.stepArgs.append(kwargs[0])
            elif cond == "exit":
                self.conds.append(cond)
                self.stepArgs.append(None)
            elif cond == "check send":
                self.conds.append(cond)
                self.stepArgs.append(None)
            else: raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Invalid cond given to AddStep")
        except Exception as error: return error
        else: return 0
    
    def Execute(self):
        """Executes the steps specified in self.conds.

        Executes each step in self.conds using the parameters in self.stepArgs.
        Loops until the user has pressed the escape key, or an exit step has
        been reached.  Once one of these two events has occurred, Cleanup() is
        called.

        Raises:
            Exception: A certain required member was not set.
        """

        if self.driverAddress == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify driver address.")
        if self.exitCondition == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify exit condition.")
        if self.exitConditionSearchType == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify exit condition search type.")
        if self.exitConditionSearchElement == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify exit condition search element.")
        if self.to == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify send address.")
        if self.fromm == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify from identifier.")
        if self.email == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify send email.")
        if self.emailPassword == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify send password.")
        if self.sendMessage == None:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! Must specify send message.")
        if len(self.conds) == 0:
            raise Exception(MAGENTA + "LOOKOUT >" + RESET + " Unable to execute! There must be at least one step.")
        running = True
        while running:
            self.DriverLoad(self.driverAddress)
            for i in range(len(self.conds)):
                if msvcrt.kbhit() and msvcrt.getch()[0] == 27:
                    print(MAGENTA + "LOOKOUT >" + RESET + " Keyboard interrupt detected")
                    running = False
                    break
                cond = self.conds[i]
                stepArg = self.stepArgs[i]
                if cond == "wait for seconds":
                    startMS = time.time() * 1000.0
                    current = startMS
                    while (current - startMS) < (stepArg * 1000):
                        progress = (50 * (current - startMS)) / (stepArg * 1000)
                        print(MAGENTA + "LOOKOUT >" + RESET + " Sleeping for " + str(int((stepArg * 1000 - (current - startMS)) / 1000) + 1) + " seconds... [" + ("#" * round(progress)) + (" " * (50 - round(progress))) + "]\r", end="")
                        current = time.time() * 1000.0
                    print()
                elif cond == "exit":
                    print(MAGENTA + "LOOKOUT >" + RESET + " Exiting")
                    running = False
                    self.Cleanup()
                elif cond == "login":
                    print(MAGENTA + "LOOKOUT >" + RESET + " Logging in with the following details:\nUsername: " + stepArg[2] + "\nPassword: " + stepArg[5])
                    user = self.FindFirstElement(stepArg[0], stepArg[1])
                    if user != None:
                        try: user.send_keys(stepArg[2])
                        except Exception as error: print(GREEN + "DRIVER >" + RESET + " " + error)
                        password = self.FindFirstElement(stepArg[3], stepArg[4])
                        if password != None:
                            try: password.send_keys(stepArg[5])
                            except Exception as error: print(GREEN + "DRIVER >" + RESET + " " + error)
                            submit = self.FindFirstElement(stepArg[6], stepArg[7])
                            if submit != None:
                                try:
                                    submit.click()
                                    print(MAGENTA + "LOOKOUT >" + RESET + " Logged in")
                                except Exception as error: print(GREEN + "DRIVER >" + RESET + " " + error)
                            else: print (MAGENTA + "LOOKOUT >" + RESET + " Unable to login becuase submit element was not found")
                        else: print(MAGENTA + "LOOKOUT >" + RESET + " Unable to login because password element was not found.")
                    else: print(MAGENTA + "LOOKOUT >" + RESET + " Unable to login because username element was not found.")
                elif cond == "check send":
                    shouldExit = False
                    element = self.FindFirstElement(self.exitConditionSearchType, self.exitConditionSearchElement)
                    if self.exitCondition == "found": shouldExit = element != None
                    elif self.exitCondition == "not found": shouldExit = element == None
                    #TODO: add changed
                    if shouldExit:
                        print(MAGENTA + "LOOKOUT >" + RESET + " Exit condition met")
                        running = False
                        self.SendUpdate()
                    else: print(MAGENTA + "LOOKOUT >" + RESET + " Exit condition not met")
            if running == False:
                print(MAGENTA + "LOOKOUT >" + RESET + " Exiting")
                self.Cleanup()