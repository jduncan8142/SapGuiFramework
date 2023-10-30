import os
from SapGuiFramework.Flow.Data import load_case_from_json_file

def test_load_case_from_json_file():
    # given
    data_file = os.path.join(os.path.dirname(__file__), 'data', 'test_case.json')
    expected_case = Case(
        name='Test Case',
        steps=[
            Step(
                description='Step 1',
                actions=['Action 1.1', 'Action 1.2'],
                expected_results=['Expected Result 1.1', 'Expected Result 1.2']
            ),
            Step(
                description='Step 2',
                actions=['Action 2.1', 'Action 2.2'],
                expected_results=['Expected Result 2.1', 'Expected Result 2.2']
            )
        ]
    )

    # when
    actual_case = load_case_from_json_file(data_file)

    # then
    assert actual_case == expected_case