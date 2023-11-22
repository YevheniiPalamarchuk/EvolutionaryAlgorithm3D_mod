using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace EvolutionaryAlgorithm3D
{
    class Program
    {
        static void Main(string[] args)
        {
            // Outer loop for running the algorithm 100 times
            for (int run = 1; run <= 100; run++)
            {
                Console.WriteLine($"Run {run}:");

                // Inner loop for 1000 generations
                for (int iteration = 1; iteration <= 1000; iteration++)
                {
                    Random rnd = new Random();
                    List<(double[], double)> population = new List<(double[], double)>();

                    // Area
                    double minX = -30.0;
                    double maxX = 30.0;
                    double minY = -30.0;
                    double maxY = 30.0;
                    double minZ = -30.0;
                    double maxZ = 30.0;

                    // Population size
                    int populationSize = 100;

                    // Evolutionary Algorithm Parameters
                    int generations = 100;
                    double mutationRate = 0.1;

                    // Initialize population
                    for (int i = 0; i < populationSize; i++)
                    {
                        double[] individual = {
                            rnd.NextDouble() * (maxX - minX) + minX,
                            rnd.NextDouble() * (maxY - minY) + minY,
                            rnd.NextDouble() * (maxZ - minZ) + minZ
                        };

                        double distance = CalculateDistance(individual, new double[] { 0.0, 0.0, 0.0 });
                        population.Add((individual, distance));
                    }

                    for (int generation = 0; generation < generations; generation++)
                    {
                        // Selection
                        List<(double[], double)> selectedPopulation = TournamentSelection(population, 5);

                        // Recombination (Crossover)
                        List<(double[], double)> offspring = Recombination(selectedPopulation);

                        // Mutation
                        offspring = Mutation(offspring, mutationRate, minX, maxX, minY, maxY, minZ, maxZ);

                        // Reproduction (Replacement)
                        population = Reproduction(population, offspring);

                        // Print best individual in the current generation
                        var bestIndividual = GetBestIndividual(population);
                        Console.WriteLine($"Generation {generation + 1}, Best Distance: {bestIndividual.Item2}");
                    }

                    // Generate Excel File with timestamp in the name
                    GenerateExcelFile(population, $"Run_{run}_Iteration_{iteration}_EvolutionaryAlgorithmPoints_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                }

                // You can add a line here to separate the output of each run
                Console.WriteLine(new string('-', 40));
            }

            Console.WriteLine("All runs completed.");
        }

        // Additional Functions

        static double CalculateDistance(double[] coordinates1, double[] aim)
        {
            return Math.Sqrt(Math.Pow(coordinates1[0] - aim[0], 2) + Math.Pow(coordinates1[1] - aim[1], 2) + Math.Pow(coordinates1[2] - aim[2], 2));
        }

        static List<(double[], double)> TournamentSelection(List<(double[], double)> population, int tournamentSize)
        {
            Random rnd = new Random();
            List<(double[], double)> selectedPopulation = new List<(double[], double)>();

            for (int i = 0; i < population.Count; i++)
            {
                // Randomly select individuals for the tournament
                List<(double[], double)> tournament = new List<(double[], double)>();
                for (int j = 0; j < tournamentSize; j++)
                {
                    int randomIndex = rnd.Next(population.Count);
                    tournament.Add(population[randomIndex]);
                }

                // Select the best individual from the tournament
                var bestIndividual = GetBestIndividual(tournament);
                selectedPopulation.Add(bestIndividual);
            }

            return selectedPopulation;
        }

        static List<(double[], double)> Recombination(List<(double[], double)> selectedPopulation)
        {
            Random rnd = new Random();
            List<(double[], double)> offspring = new List<(double[], double)>();

            for (int i = 0; i < selectedPopulation.Count; i += 2)
            {
                // Two-point crossover
                int crossoverPoint1 = rnd.Next(selectedPopulation[i].Item1.Length);
                int crossoverPoint2 = rnd.Next(crossoverPoint1, selectedPopulation[i].Item1.Length);

                double[] child1 = new double[selectedPopulation[i].Item1.Length];
                double[] child2 = new double[selectedPopulation[i].Item1.Length];

                for (int j = 0; j < crossoverPoint1; j++)
                {
                    child1[j] = selectedPopulation[i].Item1[j];
                    child2[j] = selectedPopulation[i + 1].Item1[j];
                }

                for (int j = crossoverPoint1; j < crossoverPoint2; j++)
                {
                    child1[j] = selectedPopulation[i + 1].Item1[j];
                    child2[j] = selectedPopulation[i].Item1[j];
                }

                for (int j = crossoverPoint2; j < selectedPopulation[i].Item1.Length; j++)
                {
                    child1[j] = selectedPopulation[i].Item1[j];
                    child2[j] = selectedPopulation[i + 1].Item1[j];
                }

                offspring.Add((child1, CalculateDistance(child1, new double[] { 0.0, 0.0, 0.0 })));
                offspring.Add((child2, CalculateDistance(child2, new double[] { 0.0, 0.0, 0.0 })));
            }

            return offspring;
        }

        static List<(double[], double)> Mutation(List<(double[], double)> offspring, double mutationRate, double minX, double maxX, double minY, double maxY, double minZ, double maxZ)
        {
            Random rnd = new Random();
            for (int i = 0; i < offspring.Count; i++)
            {
                if (rnd.NextDouble() < mutationRate)
                {
                    int mutationPoint = rnd.Next(offspring[i].Item1.Length);
                    double mutationValue = rnd.NextDouble() * (maxX - minX) + minX;
                    offspring[i].Item1[mutationPoint] = mutationValue;
                    offspring[i] = (offspring[i].Item1, CalculateDistance(offspring[i].Item1, new double[] { 0.0, 0.0, 0.0 }));
                }
            }

            return offspring;
        }

        static List<(double[], double)> Reproduction(List<(double[], double)> population, List<(double[], double)> offspring)
        {
            // Combine the population and offspring, then select the top individuals
            List<(double[], double)> combinedPopulation = new List<(double[], double)>(population);
            combinedPopulation.AddRange(offspring);

            return GetTopIndividuals(combinedPopulation, population.Count);
        }

        static (double[], double) GetBestIndividual(List<(double[], double)> individuals)
        {
            // Return the individual with the lowest distance
            double minDistance = double.MaxValue;
            (double[], double) bestIndividual = (null, 0.0);

            foreach (var individual in individuals)
            {
                if (individual.Item2 < minDistance)
                {
                    minDistance = individual.Item2;
                    bestIndividual = individual;
                }
            }

            return bestIndividual;
        }

        static List<(double[], double)> GetTopIndividuals(List<(double[], double)> individuals, int topCount)
        {
            // Return the top individuals based on distance
            individuals.Sort((a, b) => a.Item2.CompareTo(b.Item2));
            return individuals.GetRange(0, topCount);
        }

        static void GenerateExcelFile(List<(double[], double)> points, string fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Points");

                // Adding headers
                worksheet.Cells[1, 1].Value = "X";
                worksheet.Cells[1, 2].Value = "Y";
                worksheet.Cells[1, 3].Value = "Z";
                worksheet.Cells[1, 4].Value = "Distance";

                // Adding data
                for (int i = 0; i < points.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = points[i].Item1[0];
                    worksheet.Cells[i + 2, 2].Value = points[i].Item1[1];
                    worksheet.Cells[i + 2, 3].Value = points[i].Item1[2];
                    worksheet.Cells[i + 2, 4].Value = points[i].Item2;
                }

                package.Save();
            }
        }
    }
}
