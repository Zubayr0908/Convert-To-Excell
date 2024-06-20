import pandas as pd
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from config import DB_HOST, DB_USER, DB_PASSWORD, DB_NAME
from models import User, Test, Question, Participation, Subject, Role, async_session, engine

async def fetch_data_to_excel(filename):
    async with async_session() as session:
        async with session.begin():
            # Example query to fetch data from multiple tables
            query = (
                select(User, Test, Question)
                .join(Test, User.id == Test.ownerID)
                .join(Question, Test.testID == Question.testID)
            )
            result = await session.execute(query)
            data = result.fetchall()

            # Convert SQL result to pandas DataFrame
            df = pd.DataFrame(data, columns=['User', 'Test', 'Question'])

            # Write DataFrame to Excel
            df.to_excel(filename, index=False, engine='openpyxl')

async def main():
    await fetch_data_to_excel('output.xlsx')

if __name__ == '__main__':
    import asyncio
    loop = asyncio.get_event_loop()
    loop.run_until_complete(main())

