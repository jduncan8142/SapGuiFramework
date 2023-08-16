# Version: 0.1.5
1. Update Core.Framework.Session to accept case parameter to allow user to pass in specific case object. 
2. Remove redundant logger creation.
3. Parse Flow.Data.Case.Steps during Session initialization. 
4. Move load_case_from_json_file from Core.Framework.Session to Flow.Data
5. Updated load_case_from_json_file function to check for values in the following order:
    a. The json being loaded.
    b. An environment variable
    c. A generic default value
6. Bump version from 0.1.4 to 0.1.5. 
7. Add CaseTypes Enum class to Flow.Data. 
    CaseTypes can be "GUI" or "WEB"
8. Add CaseType attribute to Flow.Data.Case. 
9. Updated types in Flow.Data.Case to reflect actual expected values.
10. Added load_case_from_excel_file to Flow.Data. 
    This is not yet implemented and will raise a NotImplementedError
11. Split load_case_from_json_file into load_case_from_json_file and load_case functions. 
    This splits the Case loading from the json parsing logic so load_case_from_excel_file can 
    use the same load_case function. 
12. Added WEB attributes to Flow.Data.Case. 
13. Added run_steps to Core.Framework.Session. 
14. Added run function to execute Actions within Flow.Actions.Step. 
15. Reorg'ed Core.Framework.Session's __init__ and __post_init__ functions. 