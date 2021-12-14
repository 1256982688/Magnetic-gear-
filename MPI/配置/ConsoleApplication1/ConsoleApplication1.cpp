#include<mpi.h>

#include <stdio.h>
#include <stdlib.h>
#pragma warning(disable : 4996)

int main(int* argc, char* argv[])
{
    int myid;
    MPI_Init(NULL, &argv);
    MPI_Comm_rank(MPI_COMM_WORLD, &myid);
    if (myid == 0)
    {
        printf("Hello!");
    }
    if (myid == 1)
    {
        printf("Hi!");
    }
    MPI_Finalize();
}